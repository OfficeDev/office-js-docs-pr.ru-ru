---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: a71d3869d5dbf91db7823118a8d0409699e17cd5
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629225"
---
# <a name="item"></a><span data-ttu-id="4bbe3-102">item</span><span class="sxs-lookup"><span data-stu-id="4bbe3-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4bbe3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4bbe3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4bbe3-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-mailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-mailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-106">Requirements</span></span>

|<span data-ttu-id="4bbe3-107">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-107">Requirement</span></span>|<span data-ttu-id="4bbe3-108">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-110">1.0</span></span>|
|[<span data-ttu-id="4bbe3-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4bbe3-112">Restricted</span></span>|
|[<span data-ttu-id="4bbe3-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-114">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4bbe3-115">Свойства</span><span class="sxs-lookup"><span data-stu-id="4bbe3-115">Properties</span></span>

| <span data-ttu-id="4bbe3-116">Свойство</span><span class="sxs-lookup"><span data-stu-id="4bbe3-116">Property</span></span> | <span data-ttu-id="4bbe3-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bbe3-117">Minimum</span></span><br><span data-ttu-id="4bbe3-118">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-118">permission level</span></span> | <span data-ttu-id="4bbe3-119">Способов</span><span class="sxs-lookup"><span data-stu-id="4bbe3-119">Modes</span></span> | <span data-ttu-id="4bbe3-120">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="4bbe3-120">Return type</span></span> | <span data-ttu-id="4bbe3-121">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bbe3-121">Minimum</span></span><br><span data-ttu-id="4bbe3-122">набор требований</span><span class="sxs-lookup"><span data-stu-id="4bbe3-122">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="4bbe3-123">attachments</span><span class="sxs-lookup"><span data-stu-id="4bbe3-123">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="4bbe3-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-124">ReadItem</span></span> | <span data-ttu-id="4bbe3-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-125">Read</span></span> | <span data-ttu-id="4bbe3-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span> | <span data-ttu-id="4bbe3-127">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-127">1.0</span></span> |
| [<span data-ttu-id="4bbe3-128">bcc</span><span class="sxs-lookup"><span data-stu-id="4bbe3-128">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="4bbe3-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-129">ReadItem</span></span> | <span data-ttu-id="4bbe3-130">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-130">Message Compose</span></span> | [<span data-ttu-id="4bbe3-131">Получатели</span><span class="sxs-lookup"><span data-stu-id="4bbe3-131">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="4bbe3-132">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-132">1.1</span></span> |
| [<span data-ttu-id="4bbe3-133">body</span><span class="sxs-lookup"><span data-stu-id="4bbe3-133">body</span></span>](#body-body) | <span data-ttu-id="4bbe3-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-134">ReadItem</span></span> | <span data-ttu-id="4bbe3-135">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-135">Compose</span></span> | [<span data-ttu-id="4bbe3-136">Body</span><span class="sxs-lookup"><span data-stu-id="4bbe3-136">Body</span></span>](/javascript/api/outlook/office.body) | <span data-ttu-id="4bbe3-137">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-137">1.1</span></span> |
| | | <span data-ttu-id="4bbe3-138">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-138">Read</span></span> | | |
| [<span data-ttu-id="4bbe3-139">categories</span><span class="sxs-lookup"><span data-stu-id="4bbe3-139">categories</span></span>](#categories-categories) | <span data-ttu-id="4bbe3-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-140">ReadItem</span></span> | <span data-ttu-id="4bbe3-141">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-141">Compose</span></span> | [<span data-ttu-id="4bbe3-142">Categories</span><span class="sxs-lookup"><span data-stu-id="4bbe3-142">Categories</span></span>](/javascript/api/outlook/office.categories) | <span data-ttu-id="4bbe3-143">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-143">Preview</span></span> |
| | | <span data-ttu-id="4bbe3-144">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-144">Read</span></span> | | |
| [<span data-ttu-id="4bbe3-145">cc</span><span class="sxs-lookup"><span data-stu-id="4bbe3-145">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4bbe3-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-146">ReadItem</span></span> | <span data-ttu-id="4bbe3-147">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-147">Message Compose</span></span> | [<span data-ttu-id="4bbe3-148">Получатели</span><span class="sxs-lookup"><span data-stu-id="4bbe3-148">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="4bbe3-149">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-149">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-150">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-150">Message Read</span></span> | <span data-ttu-id="4bbe3-151">Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-151">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="4bbe3-152">conversationId</span><span class="sxs-lookup"><span data-stu-id="4bbe3-152">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4bbe3-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-153">ReadItem</span></span> | <span data-ttu-id="4bbe3-154">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-154">Message Compose</span></span> | <span data-ttu-id="4bbe3-155">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-155">String</span></span> | <span data-ttu-id="4bbe3-156">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-156">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-157">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-157">Message Read</span></span> | | |
| [<span data-ttu-id="4bbe3-158">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4bbe3-158">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4bbe3-159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-159">ReadItem</span></span> | <span data-ttu-id="4bbe3-160">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-160">Read</span></span> | <span data-ttu-id="4bbe3-161">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-161">Date</span></span> | <span data-ttu-id="4bbe3-162">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-162">1.0</span></span> |
| [<span data-ttu-id="4bbe3-163">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4bbe3-163">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4bbe3-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-164">ReadItem</span></span> | <span data-ttu-id="4bbe3-165">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-165">Read</span></span> | <span data-ttu-id="4bbe3-166">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-166">Date</span></span> | <span data-ttu-id="4bbe3-167">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-167">1.0</span></span> |
| [<span data-ttu-id="4bbe3-168">end</span><span class="sxs-lookup"><span data-stu-id="4bbe3-168">end</span></span>](#end-datetime) | <span data-ttu-id="4bbe3-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-169">ReadItem</span></span> | <span data-ttu-id="4bbe3-170">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-170">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-171">Time</span><span class="sxs-lookup"><span data-stu-id="4bbe3-171">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="4bbe3-172">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-172">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-173">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-173">Appointment Attendee</span></span> | <span data-ttu-id="4bbe3-174">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-174">Date</span></span> | |
| | | <span data-ttu-id="4bbe3-175">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-175">Message Read</span></span><br><span data-ttu-id="4bbe3-176">(Приглашение на собрание)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-176">(Meeting Request)</span></span> | <span data-ttu-id="4bbe3-177">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-177">Date</span></span> | |
| [<span data-ttu-id="4bbe3-178">енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="4bbe3-178">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="4bbe3-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-179">ReadItem</span></span> | <span data-ttu-id="4bbe3-180">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-180">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-181">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="4bbe3-181">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation) | <span data-ttu-id="4bbe3-182">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-182">Preview</span></span> |
| | | <span data-ttu-id="4bbe3-183">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-183">Appointment Attendee</span></span> | | |
| [<span data-ttu-id="4bbe3-184">from</span><span class="sxs-lookup"><span data-stu-id="4bbe3-184">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="4bbe3-185">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-185">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-186">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-186">Message Compose</span></span> | [<span data-ttu-id="4bbe3-187">From</span><span class="sxs-lookup"><span data-stu-id="4bbe3-187">From</span></span>](/javascript/api/outlook/office.from) | <span data-ttu-id="4bbe3-188">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-188">1.7</span></span> |
| | <span data-ttu-id="4bbe3-189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-189">ReadItem</span></span> | <span data-ttu-id="4bbe3-190">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-190">Message Read</span></span> | [<span data-ttu-id="4bbe3-191">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bbe3-191">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="4bbe3-192">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-192">1.0</span></span> |
| [<span data-ttu-id="4bbe3-193">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-193">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="4bbe3-194">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-194">ReadItem</span></span> | <span data-ttu-id="4bbe3-195">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-195">Message Compose</span></span> | [<span data-ttu-id="4bbe3-196">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="4bbe3-196">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders) | <span data-ttu-id="4bbe3-197">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-197">Preview</span></span> |
| [<span data-ttu-id="4bbe3-198">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4bbe3-198">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4bbe3-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-199">ReadItem</span></span> | <span data-ttu-id="4bbe3-200">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-200">Message Read</span></span> | <span data-ttu-id="4bbe3-201">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-201">String</span></span> | <span data-ttu-id="4bbe3-202">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-202">1.0</span></span> |
| [<span data-ttu-id="4bbe3-203">itemClass</span><span class="sxs-lookup"><span data-stu-id="4bbe3-203">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4bbe3-204">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-204">ReadItem</span></span> | <span data-ttu-id="4bbe3-205">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-205">Read</span></span> | <span data-ttu-id="4bbe3-206">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-206">String</span></span> | <span data-ttu-id="4bbe3-207">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-207">1.0</span></span> |
| [<span data-ttu-id="4bbe3-208">itemId</span><span class="sxs-lookup"><span data-stu-id="4bbe3-208">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4bbe3-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-209">ReadItem</span></span> | <span data-ttu-id="4bbe3-210">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-210">Read</span></span> | <span data-ttu-id="4bbe3-211">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-211">String</span></span> | <span data-ttu-id="4bbe3-212">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-212">1.0</span></span> |
| [<span data-ttu-id="4bbe3-213">itemType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-213">itemType</span></span>](#itemtype-mailboxenumsitemtype) | <span data-ttu-id="4bbe3-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-214">ReadItem</span></span> | <span data-ttu-id="4bbe3-215">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-215">Compose</span></span> | [<span data-ttu-id="4bbe3-216">MailboxEnums. ItemType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-216">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype) | <span data-ttu-id="4bbe3-217">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-217">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-218">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-218">Read</span></span> | | |
| [<span data-ttu-id="4bbe3-219">location</span><span class="sxs-lookup"><span data-stu-id="4bbe3-219">location</span></span>](#location-stringlocation) | <span data-ttu-id="4bbe3-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-220">ReadItem</span></span> | <span data-ttu-id="4bbe3-221">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-221">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-222">Location</span><span class="sxs-lookup"><span data-stu-id="4bbe3-222">Location</span></span>](/javascript/api/outlook/office.location) | <span data-ttu-id="4bbe3-223">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-223">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-224">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-224">Appointment Attendee</span></span> | <span data-ttu-id="4bbe3-225">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-225">String</span></span> | |
| | | <span data-ttu-id="4bbe3-226">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-226">Message Read</span></span><br><span data-ttu-id="4bbe3-227">(Приглашение на собрание)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-227">(Meeting Request)</span></span> | <span data-ttu-id="4bbe3-228">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-228">String</span></span> | |
| [<span data-ttu-id="4bbe3-229">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4bbe3-229">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4bbe3-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-230">ReadItem</span></span> | <span data-ttu-id="4bbe3-231">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-231">Read</span></span> | <span data-ttu-id="4bbe3-232">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-232">String</span></span> | <span data-ttu-id="4bbe3-233">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-233">1.0</span></span> |
| [<span data-ttu-id="4bbe3-234">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4bbe3-234">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="4bbe3-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-235">ReadItem</span></span> | <span data-ttu-id="4bbe3-236">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-236">Message Compose</span></span> | [<span data-ttu-id="4bbe3-237">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4bbe3-237">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages) | <span data-ttu-id="4bbe3-238">1.3</span><span class="sxs-lookup"><span data-stu-id="4bbe3-238">1.3</span></span> |
| | <span data-ttu-id="4bbe3-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-239">ReadItem</span></span> | <span data-ttu-id="4bbe3-240">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-240">Message Read</span></span> | | |
| [<span data-ttu-id="4bbe3-241">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4bbe3-241">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4bbe3-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-242">ReadItem</span></span> | <span data-ttu-id="4bbe3-243">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-243">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-244">Получатели</span><span class="sxs-lookup"><span data-stu-id="4bbe3-244">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="4bbe3-245">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-245">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-246">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-246">Appointment Attendee</span></span> | <span data-ttu-id="4bbe3-247">Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-247">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="4bbe3-248">organizer</span><span class="sxs-lookup"><span data-stu-id="4bbe3-248">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="4bbe3-249">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-249">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-250">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-250">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-251">Organizer</span><span class="sxs-lookup"><span data-stu-id="4bbe3-251">Organizer</span></span>](/javascript/api/outlook/office.organizer) | <span data-ttu-id="4bbe3-252">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-252">1.7</span></span> |
| | <span data-ttu-id="4bbe3-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-253">ReadItem</span></span> | <span data-ttu-id="4bbe3-254">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-254">Appointment Attendee</span></span> | [<span data-ttu-id="4bbe3-255">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bbe3-255">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="4bbe3-256">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-256">1.0</span></span> |
| [<span data-ttu-id="4bbe3-257">recurrence</span><span class="sxs-lookup"><span data-stu-id="4bbe3-257">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="4bbe3-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-258">ReadItem</span></span> | <span data-ttu-id="4bbe3-259">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-259">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-260">Повторения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-260">Recurrence</span></span>](/javascript/api/outlook/office.recurrence) | <span data-ttu-id="4bbe3-261">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-261">1.7</span></span> |
| | | <span data-ttu-id="4bbe3-262">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-262">Appointment Attendee</span></span> | | |
| | | <span data-ttu-id="4bbe3-263">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-263">Message Read</span></span><br><span data-ttu-id="4bbe3-264">(Приглашение на собрание)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-264">(Meeting Request)</span></span> | | |
| [<span data-ttu-id="4bbe3-265">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4bbe3-265">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4bbe3-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-266">ReadItem</span></span> | <span data-ttu-id="4bbe3-267">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-267">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-268">Получатели</span><span class="sxs-lookup"><span data-stu-id="4bbe3-268">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="4bbe3-269">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-269">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-270">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-270">Appointment Attendee</span></span> | <span data-ttu-id="4bbe3-271">Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-271">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="4bbe3-272">sender</span><span class="sxs-lookup"><span data-stu-id="4bbe3-272">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="4bbe3-273">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-273">ReadItem</span></span> | <span data-ttu-id="4bbe3-274">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-274">Message Read</span></span> | [<span data-ttu-id="4bbe3-275">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bbe3-275">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="4bbe3-276">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-276">1.0</span></span> |
| [<span data-ttu-id="4bbe3-277">seriesId</span><span class="sxs-lookup"><span data-stu-id="4bbe3-277">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="4bbe3-278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-278">ReadItem</span></span> | <span data-ttu-id="4bbe3-279">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-279">Compose</span></span> | <span data-ttu-id="4bbe3-280">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-280">String</span></span> | <span data-ttu-id="4bbe3-281">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-281">1.7</span></span> |
| | | <span data-ttu-id="4bbe3-282">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-282">Read</span></span> | | |
| [<span data-ttu-id="4bbe3-283">start</span><span class="sxs-lookup"><span data-stu-id="4bbe3-283">start</span></span>](#start-datetime) | <span data-ttu-id="4bbe3-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-284">ReadItem</span></span> | <span data-ttu-id="4bbe3-285">Организатор встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-285">Appointment Organizer</span></span> | [<span data-ttu-id="4bbe3-286">Time</span><span class="sxs-lookup"><span data-stu-id="4bbe3-286">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="4bbe3-287">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-287">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-288">Участник встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-288">Appointment Attendee</span></span> | <span data-ttu-id="4bbe3-289">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-289">Date</span></span> | |
| | | <span data-ttu-id="4bbe3-290">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-290">Message Read</span></span><br><span data-ttu-id="4bbe3-291">(Приглашение на собрание)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-291">(Meeting Request)</span></span> | <span data-ttu-id="4bbe3-292">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-292">Date</span></span> | |
| [<span data-ttu-id="4bbe3-293">subject</span><span class="sxs-lookup"><span data-stu-id="4bbe3-293">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="4bbe3-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-294">ReadItem</span></span> | <span data-ttu-id="4bbe3-295">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-295">Compose</span></span> | [<span data-ttu-id="4bbe3-296">Subject</span><span class="sxs-lookup"><span data-stu-id="4bbe3-296">Subject</span></span>](/javascript/api/outlook/office.subject) | <span data-ttu-id="4bbe3-297">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-297">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-298">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-298">Read</span></span> | <span data-ttu-id="4bbe3-299">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-299">String</span></span> | |
| [<span data-ttu-id="4bbe3-300">to</span><span class="sxs-lookup"><span data-stu-id="4bbe3-300">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4bbe3-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-301">ReadItem</span></span> | <span data-ttu-id="4bbe3-302">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-302">Message Compose</span></span> | [<span data-ttu-id="4bbe3-303">Получатели</span><span class="sxs-lookup"><span data-stu-id="4bbe3-303">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="4bbe3-304">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-304">1.0</span></span> |
| | | <span data-ttu-id="4bbe3-305">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-305">Message Read</span></span> | <span data-ttu-id="4bbe3-306">Массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-306">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |

##### <a name="methods"></a><span data-ttu-id="4bbe3-307">Методы</span><span class="sxs-lookup"><span data-stu-id="4bbe3-307">Methods</span></span>

| <span data-ttu-id="4bbe3-308">Метод</span><span class="sxs-lookup"><span data-stu-id="4bbe3-308">Method</span></span> | <span data-ttu-id="4bbe3-309">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bbe3-309">Minimum</span></span><br><span data-ttu-id="4bbe3-310">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-310">permission level</span></span> | <span data-ttu-id="4bbe3-311">Способов</span><span class="sxs-lookup"><span data-stu-id="4bbe3-311">Modes</span></span> | <span data-ttu-id="4bbe3-312">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bbe3-312">Minimum</span></span><br><span data-ttu-id="4bbe3-313">набор требований</span><span class="sxs-lookup"><span data-stu-id="4bbe3-313">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="4bbe3-314">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-314">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4bbe3-315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-315">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-316">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-316">Compose</span></span> | <span data-ttu-id="4bbe3-317">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-317">1.1</span></span> |
| [<span data-ttu-id="4bbe3-318">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="4bbe3-318">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="4bbe3-319">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-319">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-320">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-320">Compose</span></span> | <span data-ttu-id="4bbe3-321">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-321">Preview</span></span> |
| [<span data-ttu-id="4bbe3-322">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-322">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4bbe3-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-323">ReadItem</span></span> | <span data-ttu-id="4bbe3-324">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-324">Compose</span></span><br><span data-ttu-id="4bbe3-325">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-325">Read</span></span> | <span data-ttu-id="4bbe3-326">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-326">1.7</span></span> |
| [<span data-ttu-id="4bbe3-327">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-327">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4bbe3-328">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-328">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-329">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-329">Compose</span></span> | <span data-ttu-id="4bbe3-330">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-330">1.1</span></span> |
| [<span data-ttu-id="4bbe3-331">close</span><span class="sxs-lookup"><span data-stu-id="4bbe3-331">close</span></span>](#close) | <span data-ttu-id="4bbe3-332">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4bbe3-332">Restricted</span></span> | <span data-ttu-id="4bbe3-333">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-333">Compose</span></span> | <span data-ttu-id="4bbe3-334">1.3</span><span class="sxs-lookup"><span data-stu-id="4bbe3-334">1.3</span></span> |
| [<span data-ttu-id="4bbe3-335">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4bbe3-335">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4bbe3-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-336">ReadItem</span></span> | <span data-ttu-id="4bbe3-337">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-337">Read</span></span> | <span data-ttu-id="4bbe3-338">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-338">1.0</span></span> |
| [<span data-ttu-id="4bbe3-339">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4bbe3-339">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4bbe3-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-340">ReadItem</span></span> | <span data-ttu-id="4bbe3-341">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-341">Read</span></span> | <span data-ttu-id="4bbe3-342">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-342">1.0</span></span> |
| [<span data-ttu-id="4bbe3-343">жеталлинтернесеадерсасинк</span><span class="sxs-lookup"><span data-stu-id="4bbe3-343">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="4bbe3-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-344">ReadItem</span></span> | <span data-ttu-id="4bbe3-345">Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-345">Message Read</span></span> | <span data-ttu-id="4bbe3-346">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-346">1.8</span></span> |
| [<span data-ttu-id="4bbe3-347">жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="4bbe3-347">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="4bbe3-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-348">ReadItem</span></span> | <span data-ttu-id="4bbe3-349">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-349">Compose</span></span><br><span data-ttu-id="4bbe3-350">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-350">Read</span></span> | <span data-ttu-id="4bbe3-351">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-351">Preview</span></span> |
| [<span data-ttu-id="4bbe3-352">жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="4bbe3-352">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="4bbe3-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-353">ReadItem</span></span> | <span data-ttu-id="4bbe3-354">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-354">Compose</span></span> | <span data-ttu-id="4bbe3-355">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-355">Preview</span></span> |
| [<span data-ttu-id="4bbe3-356">getEntities</span><span class="sxs-lookup"><span data-stu-id="4bbe3-356">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="4bbe3-357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-357">ReadItem</span></span> | <span data-ttu-id="4bbe3-358">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-358">Read</span></span> | <span data-ttu-id="4bbe3-359">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-359">1.0</span></span> |
| [<span data-ttu-id="4bbe3-360">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-360">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4bbe3-361">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4bbe3-361">Restricted</span></span> | <span data-ttu-id="4bbe3-362">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-362">Read</span></span> | <span data-ttu-id="4bbe3-363">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-363">1.0</span></span> |
| [<span data-ttu-id="4bbe3-364">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4bbe3-364">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4bbe3-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-365">ReadItem</span></span> | <span data-ttu-id="4bbe3-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-366">Read</span></span> | <span data-ttu-id="4bbe3-367">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-367">1.0</span></span> |
| [<span data-ttu-id="4bbe3-368">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-368">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="4bbe3-369">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-369">ReadItem</span></span> | <span data-ttu-id="4bbe3-370">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-370">Read</span></span> | <span data-ttu-id="4bbe3-371">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-371">Preview</span></span> |
| [<span data-ttu-id="4bbe3-372">жетитемидасинк</span><span class="sxs-lookup"><span data-stu-id="4bbe3-372">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="4bbe3-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-373">ReadItem</span></span> | <span data-ttu-id="4bbe3-374">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-374">Compose</span></span> | <span data-ttu-id="4bbe3-375">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-375">Preview</span></span> |
| [<span data-ttu-id="4bbe3-376">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4bbe3-376">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4bbe3-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-377">ReadItem</span></span> | <span data-ttu-id="4bbe3-378">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-378">Read</span></span> | <span data-ttu-id="4bbe3-379">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-379">1.0</span></span> |
| [<span data-ttu-id="4bbe3-380">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4bbe3-380">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4bbe3-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-381">ReadItem</span></span> | <span data-ttu-id="4bbe3-382">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-382">Read</span></span> | <span data-ttu-id="4bbe3-383">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-383">1.0</span></span> |
| [<span data-ttu-id="4bbe3-384">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-384">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4bbe3-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-385">ReadItem</span></span> | <span data-ttu-id="4bbe3-386">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-386">Compose</span></span> | <span data-ttu-id="4bbe3-387">1.2</span><span class="sxs-lookup"><span data-stu-id="4bbe3-387">1.2</span></span> |
| [<span data-ttu-id="4bbe3-388">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="4bbe3-388">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="4bbe3-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-389">ReadItem</span></span> | <span data-ttu-id="4bbe3-390">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-390">Read</span></span> | <span data-ttu-id="4bbe3-391">1.6</span><span class="sxs-lookup"><span data-stu-id="4bbe3-391">1.6</span></span> |
| [<span data-ttu-id="4bbe3-392">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="4bbe3-392">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4bbe3-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-393">ReadItem</span></span> | <span data-ttu-id="4bbe3-394">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-394">Read</span></span> | <span data-ttu-id="4bbe3-395">1.6</span><span class="sxs-lookup"><span data-stu-id="4bbe3-395">1.6</span></span> |
| [<span data-ttu-id="4bbe3-396">жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="4bbe3-396">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="4bbe3-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-397">ReadItem</span></span> | <span data-ttu-id="4bbe3-398">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-398">Compose</span></span><br><span data-ttu-id="4bbe3-399">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-399">Read</span></span> | <span data-ttu-id="4bbe3-400">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-400">Preview</span></span> |
| [<span data-ttu-id="4bbe3-401">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-401">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4bbe3-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-402">ReadItem</span></span> | <span data-ttu-id="4bbe3-403">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-403">Compose</span></span><br><span data-ttu-id="4bbe3-404">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-404">Read</span></span> | <span data-ttu-id="4bbe3-405">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-405">1.0</span></span> |
| [<span data-ttu-id="4bbe3-406">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-406">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4bbe3-407">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-407">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-408">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-408">Compose</span></span> | <span data-ttu-id="4bbe3-409">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-409">1.1</span></span> |
| [<span data-ttu-id="4bbe3-410">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-410">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4bbe3-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-411">ReadItem</span></span> | <span data-ttu-id="4bbe3-412">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-412">Compose</span></span><br><span data-ttu-id="4bbe3-413">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-413">Read</span></span> | <span data-ttu-id="4bbe3-414">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-414">1.7</span></span> |
| [<span data-ttu-id="4bbe3-415">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-415">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4bbe3-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-416">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-417">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-417">Compose</span></span> | <span data-ttu-id="4bbe3-418">1.3</span><span class="sxs-lookup"><span data-stu-id="4bbe3-418">1.3</span></span> |
| [<span data-ttu-id="4bbe3-419">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4bbe3-419">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4bbe3-420">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-420">ReadWriteItem</span></span> | <span data-ttu-id="4bbe3-421">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-421">Compose</span></span> | <span data-ttu-id="4bbe3-422">1.2</span><span class="sxs-lookup"><span data-stu-id="4bbe3-422">1.2</span></span> |

##### <a name="events"></a><span data-ttu-id="4bbe3-423">События</span><span class="sxs-lookup"><span data-stu-id="4bbe3-423">Events</span></span>

<span data-ttu-id="4bbe3-424">Вы можете подписаться на следующие события и отписаться на них, используя [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) и [removeHandlerAsync](#removehandlerasynceventtype-options-callback) соответственно.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-424">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="4bbe3-425">Событие</span><span class="sxs-lookup"><span data-stu-id="4bbe3-425">Event</span></span> | <span data-ttu-id="4bbe3-426">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-426">Description</span></span> | <span data-ttu-id="4bbe3-427">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bbe3-427">Minimum</span></span><br><span data-ttu-id="4bbe3-428">набор требований</span><span class="sxs-lookup"><span data-stu-id="4bbe3-428">requirement set</span></span> |
|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="4bbe3-429">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-429">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4bbe3-430">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-430">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4bbe3-431">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-431">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4bbe3-432">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-432">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="4bbe3-433">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-433">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="4bbe3-434">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-434">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4bbe3-435">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-435">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4bbe3-436">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-436">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4bbe3-437">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-437">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4bbe3-438">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-438">1.7</span></span> |

### <a name="example"></a><span data-ttu-id="4bbe3-439">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-439">Example</span></span>

<span data-ttu-id="4bbe3-440">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-440">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

## <a name="property-details"></a><span data-ttu-id="4bbe3-441">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="4bbe3-441">Property details</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="4bbe3-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="4bbe3-443">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-443">Gets the item's attachments as an array.</span></span> <span data-ttu-id="4bbe3-444">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-444">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-445">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-445">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4bbe3-446">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-446">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-447">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-447">Type</span></span>

*   <span data-ttu-id="4bbe3-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-449">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-449">Requirements</span></span>

|<span data-ttu-id="4bbe3-450">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-450">Requirement</span></span>|<span data-ttu-id="4bbe3-451">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-451">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-452">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-453">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-453">1.0</span></span>|
|[<span data-ttu-id="4bbe3-454">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-455">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-456">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-457">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-457">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-458">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-458">Example</span></span>

<span data-ttu-id="4bbe3-459">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-459">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4bbe3-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4bbe3-461">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-461">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4bbe3-462">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-462">Compose mode only.</span></span>

<span data-ttu-id="4bbe3-463">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-463">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-464">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-464">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4bbe3-465">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-465">Get 500 members maximum.</span></span>
- <span data-ttu-id="4bbe3-466">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-466">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-467">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-467">Type</span></span>

*   [<span data-ttu-id="4bbe3-468">Получатели</span><span class="sxs-lookup"><span data-stu-id="4bbe3-468">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-469">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-469">Requirements</span></span>

|<span data-ttu-id="4bbe3-470">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-470">Requirement</span></span>|<span data-ttu-id="4bbe3-471">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-472">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-473">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-473">1.1</span></span>|
|[<span data-ttu-id="4bbe3-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-475">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-477">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-477">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-478">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-478">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="4bbe3-479">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-479">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="4bbe3-480">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-480">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-481">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-481">Type</span></span>

*   [<span data-ttu-id="4bbe3-482">Body</span><span class="sxs-lookup"><span data-stu-id="4bbe3-482">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-483">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-483">Requirements</span></span>

|<span data-ttu-id="4bbe3-484">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-484">Requirement</span></span>|<span data-ttu-id="4bbe3-485">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-486">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-487">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-487">1.1</span></span>|
|[<span data-ttu-id="4bbe3-488">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-488">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-489">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-490">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-490">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-491">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-491">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-492">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-492">Example</span></span>

<span data-ttu-id="4bbe3-493">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-493">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4bbe3-494">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-494">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="4bbe3-495">Категории: [категории](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-495">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="4bbe3-496">Получает объект, предоставляющий методы для управления категориями элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-496">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-497">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-497">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-498">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-498">Type</span></span>

*   [<span data-ttu-id="4bbe3-499">Categories</span><span class="sxs-lookup"><span data-stu-id="4bbe3-499">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-500">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-500">Requirements</span></span>

|<span data-ttu-id="4bbe3-501">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-501">Requirement</span></span>|<span data-ttu-id="4bbe3-502">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-503">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-504">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-504">1.8</span></span>|
|[<span data-ttu-id="4bbe3-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-506">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-508">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-508">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-509">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-509">Example</span></span>

<span data-ttu-id="4bbe3-510">В этом примере возвращаются категории элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-510">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4bbe3-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4bbe3-512">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-512">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4bbe3-513">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-513">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-514">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-514">Read mode</span></span>

<span data-ttu-id="4bbe3-515">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-515">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="4bbe3-516">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-516">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-517">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-517">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-518">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-518">Compose mode</span></span>

<span data-ttu-id="4bbe3-519">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-519">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="4bbe3-520">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-520">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-521">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-521">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4bbe3-522">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-522">Get 500 members maximum.</span></span>
- <span data-ttu-id="4bbe3-523">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-523">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-524">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-524">Type</span></span>

*   <span data-ttu-id="4bbe3-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-526">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-526">Requirements</span></span>

|<span data-ttu-id="4bbe3-527">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-527">Requirement</span></span>|<span data-ttu-id="4bbe3-528">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-529">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-530">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-530">1.0</span></span>|
|[<span data-ttu-id="4bbe3-531">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-531">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-532">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-533">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-533">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-534">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-534">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="4bbe3-535">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-535">(nullable) conversationId: String</span></span>

<span data-ttu-id="4bbe3-536">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-536">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4bbe3-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4bbe3-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-541">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-541">Type</span></span>

*   <span data-ttu-id="4bbe3-542">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-542">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-543">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-543">Requirements</span></span>

|<span data-ttu-id="4bbe3-544">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-544">Requirement</span></span>|<span data-ttu-id="4bbe3-545">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-546">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-547">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-547">1.0</span></span>|
|[<span data-ttu-id="4bbe3-548">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-549">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-550">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-551">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-552">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-552">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="4bbe3-553">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="4bbe3-553">dateTimeCreated: Date</span></span>

<span data-ttu-id="4bbe3-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-556">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-556">Type</span></span>

*   <span data-ttu-id="4bbe3-557">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-557">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-558">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-558">Requirements</span></span>

|<span data-ttu-id="4bbe3-559">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-559">Requirement</span></span>|<span data-ttu-id="4bbe3-560">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-561">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-562">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-562">1.0</span></span>|
|[<span data-ttu-id="4bbe3-563">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-564">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-565">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-566">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-566">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-567">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-567">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="4bbe3-568">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="4bbe3-568">dateTimeModified: Date</span></span>

<span data-ttu-id="4bbe3-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-571">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-571">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-572">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-572">Type</span></span>

*   <span data-ttu-id="4bbe3-573">Дата</span><span class="sxs-lookup"><span data-stu-id="4bbe3-573">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-574">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-574">Requirements</span></span>

|<span data-ttu-id="4bbe3-575">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-575">Requirement</span></span>|<span data-ttu-id="4bbe3-576">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-577">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-578">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-578">1.0</span></span>|
|[<span data-ttu-id="4bbe3-579">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-580">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-581">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-582">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-582">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-583">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-583">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="4bbe3-584">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-584">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="4bbe3-585">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-585">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4bbe3-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-588">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-588">Read mode</span></span>

<span data-ttu-id="4bbe3-589">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-589">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-590">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-590">Compose mode</span></span>

<span data-ttu-id="4bbe3-591">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-591">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4bbe3-592">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-592">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4bbe3-593">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-593">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="4bbe3-594">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-594">Type</span></span>

*   <span data-ttu-id="4bbe3-595">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-595">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-596">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-596">Requirements</span></span>

|<span data-ttu-id="4bbe3-597">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-597">Requirement</span></span>|<span data-ttu-id="4bbe3-598">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-599">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-600">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-600">1.0</span></span>|
|[<span data-ttu-id="4bbe3-601">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-602">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-603">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-604">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-604">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="4bbe3-605">Енханцедлокатион: [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-605">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="4bbe3-606">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-606">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-607">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-607">Read mode</span></span>

<span data-ttu-id="4bbe3-608">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="4bbe3-608">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-609">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-609">Compose mode</span></span>

<span data-ttu-id="4bbe3-610">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-610">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-611">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-611">Type</span></span>

*   [<span data-ttu-id="4bbe3-612">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="4bbe3-612">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-613">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-613">Requirements</span></span>

|<span data-ttu-id="4bbe3-614">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-614">Requirement</span></span>|<span data-ttu-id="4bbe3-615">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-616">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-617">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-617">1.8</span></span>|
|[<span data-ttu-id="4bbe3-618">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-619">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-620">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-621">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-622">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-622">Example</span></span>

<span data-ttu-id="4bbe3-623">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-623">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="4bbe3-624">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-624">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="4bbe3-625">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-625">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="4bbe3-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-628">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-628">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-629">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-629">Read mode</span></span>

<span data-ttu-id="4bbe3-630">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-630">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-631">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-631">Compose mode</span></span>

<span data-ttu-id="4bbe3-632">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-632">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-633">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-633">Type</span></span>

*   <span data-ttu-id="4bbe3-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-635">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-635">Requirements</span></span>

|<span data-ttu-id="4bbe3-636">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-636">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4bbe3-637">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-638">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-638">1.0</span></span>|<span data-ttu-id="4bbe3-639">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-639">1.7</span></span>|
|[<span data-ttu-id="4bbe3-640">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-640">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-641">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-641">ReadItem</span></span>|<span data-ttu-id="4bbe3-642">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-642">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-643">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-643">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-644">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-644">Read</span></span>|<span data-ttu-id="4bbe3-645">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-645">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="4bbe3-646">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-646">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="4bbe3-647">Возвращает или задает настраиваемые заголовки Интернета для сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-647">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="4bbe3-648">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-648">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-649">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-649">Type</span></span>

*   [<span data-ttu-id="4bbe3-650">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="4bbe3-650">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-651">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-651">Requirements</span></span>

|<span data-ttu-id="4bbe3-652">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-652">Requirement</span></span>|<span data-ttu-id="4bbe3-653">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-654">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-655">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-655">1.8</span></span>|
|[<span data-ttu-id="4bbe3-656">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-657">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-658">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-659">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-659">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-660">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-660">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="4bbe3-661">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-661">internetMessageId: String</span></span>

<span data-ttu-id="4bbe3-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-664">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-664">Type</span></span>

*   <span data-ttu-id="4bbe3-665">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-665">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-666">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-666">Requirements</span></span>

|<span data-ttu-id="4bbe3-667">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-667">Requirement</span></span>|<span data-ttu-id="4bbe3-668">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-669">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-670">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-670">1.0</span></span>|
|[<span data-ttu-id="4bbe3-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-672">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-674">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-674">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-675">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-675">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="4bbe3-676">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-676">itemClass: String</span></span>

<span data-ttu-id="4bbe3-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4bbe3-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="4bbe3-681">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-681">Type</span></span>|<span data-ttu-id="4bbe3-682">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-682">Description</span></span>|<span data-ttu-id="4bbe3-683">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="4bbe3-683">item class</span></span>|
|---|---|---|
|<span data-ttu-id="4bbe3-684">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="4bbe3-684">Appointment items</span></span>|<span data-ttu-id="4bbe3-685">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-685">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="4bbe3-686">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-686">Message items</span></span>|<span data-ttu-id="4bbe3-687">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-687">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="4bbe3-688">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-688">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-689">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-689">Type</span></span>

*   <span data-ttu-id="4bbe3-690">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-690">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-691">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-691">Requirements</span></span>

|<span data-ttu-id="4bbe3-692">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-692">Requirement</span></span>|<span data-ttu-id="4bbe3-693">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-694">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-695">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-695">1.0</span></span>|
|[<span data-ttu-id="4bbe3-696">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-697">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-697">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-698">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-699">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-699">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-700">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-700">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4bbe3-701">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-701">(nullable) itemId: String</span></span>

<span data-ttu-id="4bbe3-p119">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-704">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-704">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="4bbe3-705">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-705">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4bbe3-706">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-706">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4bbe3-707">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-707">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4bbe3-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-710">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-710">Type</span></span>

*   <span data-ttu-id="4bbe3-711">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-711">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-712">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-712">Requirements</span></span>

|<span data-ttu-id="4bbe3-713">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-713">Requirement</span></span>|<span data-ttu-id="4bbe3-714">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-715">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-716">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-716">1.0</span></span>|
|[<span data-ttu-id="4bbe3-717">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-718">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-719">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-720">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-720">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-721">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-721">Example</span></span>

<span data-ttu-id="4bbe3-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-mailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="4bbe3-724">itemType: [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-724">itemType: [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4bbe3-725">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-725">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4bbe3-726">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-726">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-727">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-727">Type</span></span>

*   [<span data-ttu-id="4bbe3-728">MailboxEnums. ItemType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-728">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-729">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-729">Requirements</span></span>

|<span data-ttu-id="4bbe3-730">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-730">Requirement</span></span>|<span data-ttu-id="4bbe3-731">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-731">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-732">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-732">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-733">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-733">1.0</span></span>|
|[<span data-ttu-id="4bbe3-734">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-734">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-735">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-735">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-736">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-736">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-737">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-737">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-738">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-738">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="4bbe3-739">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-739">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="4bbe3-740">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-740">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-741">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-741">Read mode</span></span>

<span data-ttu-id="4bbe3-742">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-742">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-743">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-743">Compose mode</span></span>

<span data-ttu-id="4bbe3-744">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-744">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-745">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-745">Type</span></span>

*   <span data-ttu-id="4bbe3-746">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-746">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-747">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-747">Requirements</span></span>

|<span data-ttu-id="4bbe3-748">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-748">Requirement</span></span>|<span data-ttu-id="4bbe3-749">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-750">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-751">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-751">1.0</span></span>|
|[<span data-ttu-id="4bbe3-752">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-752">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-753">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-754">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-754">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-755">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-755">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4bbe3-756">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-756">normalizedSubject: String</span></span>

<span data-ttu-id="4bbe3-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4bbe3-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-761">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-761">Type</span></span>

*   <span data-ttu-id="4bbe3-762">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-762">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-763">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-763">Requirements</span></span>

|<span data-ttu-id="4bbe3-764">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-764">Requirement</span></span>|<span data-ttu-id="4bbe3-765">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-765">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-766">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-766">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-767">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-767">1.0</span></span>|
|[<span data-ttu-id="4bbe3-768">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-768">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-769">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-769">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-770">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-770">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-771">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-771">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-772">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-772">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="4bbe3-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="4bbe3-774">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-774">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-775">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-775">Type</span></span>

*   [<span data-ttu-id="4bbe3-776">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4bbe3-776">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-777">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-777">Requirements</span></span>

|<span data-ttu-id="4bbe3-778">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-778">Requirement</span></span>|<span data-ttu-id="4bbe3-779">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-780">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-781">1.3</span><span class="sxs-lookup"><span data-stu-id="4bbe3-781">1.3</span></span>|
|[<span data-ttu-id="4bbe3-782">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-782">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-783">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-783">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-784">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-784">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-785">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-785">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-786">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-786">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4bbe3-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4bbe3-788">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-788">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4bbe3-789">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-789">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-790">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-790">Read mode</span></span>

<span data-ttu-id="4bbe3-791">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-791">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="4bbe3-792">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-792">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-793">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-793">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-794">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-794">Compose mode</span></span>

<span data-ttu-id="4bbe3-795">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-795">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="4bbe3-796">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-796">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-797">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-797">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4bbe3-798">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-798">Get 500 members maximum.</span></span>
- <span data-ttu-id="4bbe3-799">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-799">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-800">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-800">Type</span></span>

*   <span data-ttu-id="4bbe3-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-802">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-802">Requirements</span></span>

|<span data-ttu-id="4bbe3-803">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-803">Requirement</span></span>|<span data-ttu-id="4bbe3-804">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-804">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-805">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-805">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-806">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-806">1.0</span></span>|
|[<span data-ttu-id="4bbe3-807">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-807">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-808">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-808">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-809">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-809">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-810">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-810">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="4bbe3-811">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-811">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="4bbe3-812">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-812">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-813">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-813">Read mode</span></span>

<span data-ttu-id="4bbe3-814">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-814">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-815">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-815">Compose mode</span></span>

<span data-ttu-id="4bbe3-816">`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer) , который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-816">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="4bbe3-817">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-817">Type</span></span>

*   <span data-ttu-id="4bbe3-818">[](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bbe3-818">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-819">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-819">Requirements</span></span>

|<span data-ttu-id="4bbe3-820">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-820">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4bbe3-821">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-822">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-822">1.0</span></span>|<span data-ttu-id="4bbe3-823">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-823">1.7</span></span>|
|[<span data-ttu-id="4bbe3-824">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-824">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-825">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-825">ReadItem</span></span>|<span data-ttu-id="4bbe3-826">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-826">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-827">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-827">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-828">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-828">Read</span></span>|<span data-ttu-id="4bbe3-829">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-829">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="4bbe3-830">(Nullable) повторение: [повторение](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-830">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="4bbe3-831">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-831">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="4bbe3-832">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-832">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="4bbe3-833">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-833">Read and compose modes for appointment items.</span></span> <span data-ttu-id="4bbe3-834">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-834">Read mode for meeting request items.</span></span>

<span data-ttu-id="4bbe3-835">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-835">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="4bbe3-836">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-836">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="4bbe3-837">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-837">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="4bbe3-838">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-838">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="4bbe3-839">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-839">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-840">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-840">Read mode</span></span>

<span data-ttu-id="4bbe3-841">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-841">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="4bbe3-842">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-842">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-843">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-843">Compose mode</span></span>

<span data-ttu-id="4bbe3-844">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-844">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="4bbe3-845">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-845">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-846">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-846">Type</span></span>

* [<span data-ttu-id="4bbe3-847">Повторения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-847">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="4bbe3-848">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-848">Requirement</span></span>|<span data-ttu-id="4bbe3-849">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-850">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-851">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-851">1.7</span></span>|
|[<span data-ttu-id="4bbe3-852">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-853">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-854">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-855">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-855">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4bbe3-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4bbe3-857">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-857">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4bbe3-858">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-858">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-859">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-859">Read mode</span></span>

<span data-ttu-id="4bbe3-860">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-860">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="4bbe3-861">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-861">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-862">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-862">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-863">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-863">Compose mode</span></span>

<span data-ttu-id="4bbe3-864">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-864">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="4bbe3-865">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-865">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-866">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-866">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4bbe3-867">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-867">Get 500 members maximum.</span></span>
- <span data-ttu-id="4bbe3-868">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-868">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-869">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-869">Type</span></span>

*   <span data-ttu-id="4bbe3-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-871">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-871">Requirements</span></span>

|<span data-ttu-id="4bbe3-872">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-872">Requirement</span></span>|<span data-ttu-id="4bbe3-873">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-874">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-875">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-875">1.0</span></span>|
|[<span data-ttu-id="4bbe3-876">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-876">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-877">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-877">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-878">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-878">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-879">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-879">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="4bbe3-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="4bbe3-p135">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4bbe3-p136">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-885">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-885">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-886">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-886">Type</span></span>

*   [<span data-ttu-id="4bbe3-887">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4bbe3-887">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4bbe3-888">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-888">Requirements</span></span>

|<span data-ttu-id="4bbe3-889">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-889">Requirement</span></span>|<span data-ttu-id="4bbe3-890">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-891">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-892">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-892">1.0</span></span>|
|[<span data-ttu-id="4bbe3-893">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-894">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-894">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-895">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-896">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-896">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-897">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-897">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="4bbe3-898">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-898">(nullable) seriesId: String</span></span>

<span data-ttu-id="4bbe3-899">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-899">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="4bbe3-900">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-900">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="4bbe3-901">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-901">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-902">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-902">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4bbe3-903">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-903">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="4bbe3-904">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-904">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4bbe3-905">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-905">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="4bbe3-906">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-906">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="4bbe3-907">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-907">Type</span></span>

* <span data-ttu-id="4bbe3-908">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-908">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-909">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-909">Requirements</span></span>

|<span data-ttu-id="4bbe3-910">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-910">Requirement</span></span>|<span data-ttu-id="4bbe3-911">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-911">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-912">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-912">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-913">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-913">1.7</span></span>|
|[<span data-ttu-id="4bbe3-914">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-914">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-915">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-915">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-916">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-916">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-917">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-917">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-918">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-918">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="4bbe3-919">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-919">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="4bbe3-920">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-920">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4bbe3-p139">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-923">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-923">Read mode</span></span>

<span data-ttu-id="4bbe3-924">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-924">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-925">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-925">Compose mode</span></span>

<span data-ttu-id="4bbe3-926">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-926">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4bbe3-927">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-927">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4bbe3-928">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-928">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="4bbe3-929">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-929">Type</span></span>

*   <span data-ttu-id="4bbe3-930">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-930">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-931">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-931">Requirements</span></span>

|<span data-ttu-id="4bbe3-932">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-932">Requirement</span></span>|<span data-ttu-id="4bbe3-933">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-934">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-935">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-935">1.0</span></span>|
|[<span data-ttu-id="4bbe3-936">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-936">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-937">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-938">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-938">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-939">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-939">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="4bbe3-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="4bbe3-941">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-941">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4bbe3-942">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-942">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-943">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-943">Read mode</span></span>

<span data-ttu-id="4bbe3-p140">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="4bbe3-946">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-946">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-947">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-947">Compose mode</span></span>
<span data-ttu-id="4bbe3-948">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-948">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-949">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-949">Type</span></span>

*   <span data-ttu-id="4bbe3-950">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-950">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-951">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-951">Requirements</span></span>

|<span data-ttu-id="4bbe3-952">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-952">Requirement</span></span>|<span data-ttu-id="4bbe3-953">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-953">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-954">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-954">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-955">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-955">1.0</span></span>|
|[<span data-ttu-id="4bbe3-956">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-956">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-957">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-957">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-958">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-958">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-959">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-959">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="4bbe3-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="4bbe3-961">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-961">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4bbe3-962">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-962">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4bbe3-963">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4bbe3-963">Read mode</span></span>

<span data-ttu-id="4bbe3-964">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-964">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="4bbe3-965">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-965">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-966">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-966">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4bbe3-967">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4bbe3-967">Compose mode</span></span>

<span data-ttu-id="4bbe3-968">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-968">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="4bbe3-969">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-969">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4bbe3-970">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-970">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4bbe3-971">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-971">Get 500 members maximum.</span></span>
- <span data-ttu-id="4bbe3-972">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-972">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4bbe3-973">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-973">Type</span></span>

*   <span data-ttu-id="4bbe3-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-975">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-975">Requirements</span></span>

|<span data-ttu-id="4bbe3-976">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-976">Requirement</span></span>|<span data-ttu-id="4bbe3-977">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-977">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-978">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-978">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-979">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-979">1.0</span></span>|
|[<span data-ttu-id="4bbe3-980">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-980">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-981">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-981">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-982">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-982">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-983">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-983">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="4bbe3-984">Сведения о методе</span><span class="sxs-lookup"><span data-stu-id="4bbe3-984">Method details</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4bbe3-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4bbe3-986">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-986">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4bbe3-987">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-987">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4bbe3-988">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-988">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-989">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-989">Parameters</span></span>
|<span data-ttu-id="4bbe3-990">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-990">Name</span></span>|<span data-ttu-id="4bbe3-991">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-991">Type</span></span>|<span data-ttu-id="4bbe3-992">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-992">Attributes</span></span>|<span data-ttu-id="4bbe3-993">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-993">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="4bbe3-994">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-994">String</span></span>||<span data-ttu-id="4bbe3-p144">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4bbe3-997">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-997">String</span></span>||<span data-ttu-id="4bbe3-p145">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1000">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1000">Object</span></span>|<span data-ttu-id="4bbe3-1001">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1002">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1003">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1003">Object</span></span>|<span data-ttu-id="4bbe3-1004">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1005">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4bbe3-1006">Boolean</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1006">Boolean</span></span>|<span data-ttu-id="4bbe3-1007">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1008">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1008">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1009">function</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1009">function</span></span>|<span data-ttu-id="4bbe3-1010">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1011">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1011">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bbe3-1012">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1012">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4bbe3-1013">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1013">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bbe3-1014">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1014">Errors</span></span>

|<span data-ttu-id="4bbe3-1015">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1015">Error code</span></span>|<span data-ttu-id="4bbe3-1016">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1016">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4bbe3-1017">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1017">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4bbe3-1018">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1018">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4bbe3-1019">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1019">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1020">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1020">Requirements</span></span>

|<span data-ttu-id="4bbe3-1021">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1021">Requirement</span></span>|<span data-ttu-id="4bbe3-1022">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1023">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1024">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1024">1.1</span></span>|
|[<span data-ttu-id="4bbe3-1025">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1026">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1026">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-1027">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1028">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1028">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bbe3-1029">Примеры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1029">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="4bbe3-1030">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1030">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="4bbe3-1031">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1031">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4bbe3-1032">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1032">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4bbe3-1033">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1033">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="4bbe3-1034">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1034">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="4bbe3-1035">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1035">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1036">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1036">Parameters</span></span>

|<span data-ttu-id="4bbe3-1037">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1037">Name</span></span>|<span data-ttu-id="4bbe3-1038">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1038">Type</span></span>|<span data-ttu-id="4bbe3-1039">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1039">Attributes</span></span>|<span data-ttu-id="4bbe3-1040">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1040">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="4bbe3-1041">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1041">String</span></span>||<span data-ttu-id="4bbe3-1042">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1042">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="4bbe3-1043">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1043">String</span></span>||<span data-ttu-id="4bbe3-p147">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1046">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1046">Object</span></span>|<span data-ttu-id="4bbe3-1047">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1048">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1049">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1049">Object</span></span>|<span data-ttu-id="4bbe3-1050">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1051">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4bbe3-1052">Boolean</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1052">Boolean</span></span>|<span data-ttu-id="4bbe3-1053">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1054">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1054">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1055">function</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1055">function</span></span>|<span data-ttu-id="4bbe3-1056">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1057">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bbe3-1058">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1058">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4bbe3-1059">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1059">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bbe3-1060">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1060">Errors</span></span>

|<span data-ttu-id="4bbe3-1061">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1061">Error code</span></span>|<span data-ttu-id="4bbe3-1062">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1062">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4bbe3-1063">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1063">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4bbe3-1064">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1064">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4bbe3-1065">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1065">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1066">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1066">Requirements</span></span>

|<span data-ttu-id="4bbe3-1067">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1067">Requirement</span></span>|<span data-ttu-id="4bbe3-1068">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1069">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1070">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1070">1.8</span></span>|
|[<span data-ttu-id="4bbe3-1071">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1072">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1072">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-1073">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1074">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1074">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bbe3-1075">Примеры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1075">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4bbe3-1076">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1076">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4bbe3-1077">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1077">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4bbe3-1078">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1078">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1079">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1079">Parameters</span></span>

| <span data-ttu-id="4bbe3-1080">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1080">Name</span></span> | <span data-ttu-id="4bbe3-1081">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1081">Type</span></span> | <span data-ttu-id="4bbe3-1082">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1082">Attributes</span></span> | <span data-ttu-id="4bbe3-1083">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1083">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4bbe3-1084">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1084">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4bbe3-1085">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1085">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4bbe3-1086">Function</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1086">Function</span></span> || <span data-ttu-id="4bbe3-p148">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4bbe3-1090">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1090">Object</span></span> | <span data-ttu-id="4bbe3-1091">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1091">&lt;optional&gt;</span></span> | <span data-ttu-id="4bbe3-1092">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1092">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4bbe3-1093">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1093">Object</span></span> | <span data-ttu-id="4bbe3-1094">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1094">&lt;optional&gt;</span></span> | <span data-ttu-id="4bbe3-1095">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1095">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4bbe3-1096">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1096">function</span></span>| <span data-ttu-id="4bbe3-1097">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1098">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1099">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1099">Requirements</span></span>

|<span data-ttu-id="4bbe3-1100">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1100">Requirement</span></span>| <span data-ttu-id="4bbe3-1101">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1101">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1102">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1102">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bbe3-1103">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1103">1.7</span></span> |
|[<span data-ttu-id="4bbe3-1104">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1104">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bbe3-1105">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1105">ReadItem</span></span> |
|[<span data-ttu-id="4bbe3-1106">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1106">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4bbe3-1107">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1107">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4bbe3-1108">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1108">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4bbe3-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4bbe3-1110">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1110">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4bbe3-p149">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4bbe3-1114">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1114">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4bbe3-1115">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1115">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1116">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1116">Parameters</span></span>

|<span data-ttu-id="4bbe3-1117">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1117">Name</span></span>|<span data-ttu-id="4bbe3-1118">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1118">Type</span></span>|<span data-ttu-id="4bbe3-1119">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1119">Attributes</span></span>|<span data-ttu-id="4bbe3-1120">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1120">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="4bbe3-1121">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1121">String</span></span>||<span data-ttu-id="4bbe3-p150">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4bbe3-1124">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1124">String</span></span>||<span data-ttu-id="4bbe3-1125">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1125">The subject of the item to be attached.</span></span> <span data-ttu-id="4bbe3-1126">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1126">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1127">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1127">Object</span></span>|<span data-ttu-id="4bbe3-1128">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1129">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1129">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1130">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1130">Object</span></span>|<span data-ttu-id="4bbe3-1131">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1131">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1132">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1132">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1133">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1133">function</span></span>|<span data-ttu-id="4bbe3-1134">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1134">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1135">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bbe3-1136">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1136">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4bbe3-1137">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1137">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bbe3-1138">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1138">Errors</span></span>

|<span data-ttu-id="4bbe3-1139">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1139">Error code</span></span>|<span data-ttu-id="4bbe3-1140">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1140">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4bbe3-1141">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1141">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1142">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1142">Requirements</span></span>

|<span data-ttu-id="4bbe3-1143">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1143">Requirement</span></span>|<span data-ttu-id="4bbe3-1144">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1144">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1145">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1146">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1146">1.1</span></span>|
|[<span data-ttu-id="4bbe3-1147">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1148">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1148">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-1149">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1150">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-1151">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1151">Example</span></span>

<span data-ttu-id="4bbe3-1152">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1152">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="close"></a><span data-ttu-id="4bbe3-1153">close()</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1153">close()</span></span>

<span data-ttu-id="4bbe3-1154">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1154">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4bbe3-p152">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1157">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1157">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4bbe3-1158">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1158">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1159">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1159">Requirements</span></span>

|<span data-ttu-id="4bbe3-1160">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1160">Requirement</span></span>|<span data-ttu-id="4bbe3-1161">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1162">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1163">1.3</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1163">1.3</span></span>|
|[<span data-ttu-id="4bbe3-1164">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1165">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1165">Restricted</span></span>|
|[<span data-ttu-id="4bbe3-1166">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1167">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1167">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4bbe3-1168">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1168">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4bbe3-1169">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1169">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1170">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1170">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4bbe3-1171">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1171">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4bbe3-1172">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1172">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4bbe3-p153">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1176">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1176">Parameters</span></span>

|<span data-ttu-id="4bbe3-1177">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1177">Name</span></span>|<span data-ttu-id="4bbe3-1178">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1178">Type</span></span>|<span data-ttu-id="4bbe3-1179">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1179">Attributes</span></span>|<span data-ttu-id="4bbe3-1180">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1180">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4bbe3-1181">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1181">String &#124; Object</span></span>||<span data-ttu-id="4bbe3-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4bbe3-1184">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1184">**OR**</span></span><br/><span data-ttu-id="4bbe3-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4bbe3-1187">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1187">String</span></span>|<span data-ttu-id="4bbe3-1188">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1188">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4bbe3-1191">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1191">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4bbe3-1192">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1192">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1193">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1193">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4bbe3-1194">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1194">String</span></span>||<span data-ttu-id="4bbe3-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4bbe3-1197">Строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1197">String</span></span>||<span data-ttu-id="4bbe3-1198">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1198">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4bbe3-1199">Строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1199">String</span></span>||<span data-ttu-id="4bbe3-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4bbe3-1202">Логический</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1202">Boolean</span></span>||<span data-ttu-id="4bbe3-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4bbe3-1205">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1205">String</span></span>||<span data-ttu-id="4bbe3-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1209">function</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1209">function</span></span>|<span data-ttu-id="4bbe3-1210">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1210">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1211">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1212">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1212">Requirements</span></span>

|<span data-ttu-id="4bbe3-1213">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1213">Requirement</span></span>|<span data-ttu-id="4bbe3-1214">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1215">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1216">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1217">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1218">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1220">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1220">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bbe3-1221">Примеры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1221">Examples</span></span>

<span data-ttu-id="4bbe3-1222">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1222">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4bbe3-1223">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1223">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4bbe3-1224">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1224">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4bbe3-1225">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1225">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="4bbe3-1226">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1226">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="4bbe3-1227">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1227">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4bbe3-1228">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1228">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4bbe3-1229">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1229">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1230">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4bbe3-1231">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1231">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4bbe3-1232">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1232">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4bbe3-p161">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1236">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1236">Parameters</span></span>

|<span data-ttu-id="4bbe3-1237">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1237">Name</span></span>|<span data-ttu-id="4bbe3-1238">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1238">Type</span></span>|<span data-ttu-id="4bbe3-1239">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1239">Attributes</span></span>|<span data-ttu-id="4bbe3-1240">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1240">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4bbe3-1241">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1241">String &#124; Object</span></span>||<span data-ttu-id="4bbe3-p162">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4bbe3-1244">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1244">**OR**</span></span><br/><span data-ttu-id="4bbe3-p163">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4bbe3-1247">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1247">String</span></span>|<span data-ttu-id="4bbe3-1248">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1248">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-p164">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4bbe3-1251">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1251">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4bbe3-1252">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1252">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1253">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1253">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4bbe3-1254">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1254">String</span></span>||<span data-ttu-id="4bbe3-p165">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4bbe3-1257">Строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1257">String</span></span>||<span data-ttu-id="4bbe3-1258">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1258">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4bbe3-1259">Строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1259">String</span></span>||<span data-ttu-id="4bbe3-p166">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4bbe3-1262">Логический</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1262">Boolean</span></span>||<span data-ttu-id="4bbe3-p167">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4bbe3-1265">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1265">String</span></span>||<span data-ttu-id="4bbe3-p168">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1269">function</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1269">function</span></span>|<span data-ttu-id="4bbe3-1270">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1271">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1272">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1272">Requirements</span></span>

|<span data-ttu-id="4bbe3-1273">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1273">Requirement</span></span>|<span data-ttu-id="4bbe3-1274">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1274">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1275">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1275">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1276">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1276">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1277">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1277">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1278">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1279">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1279">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1280">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1280">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bbe3-1281">Примеры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1281">Examples</span></span>

<span data-ttu-id="4bbe3-1282">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1282">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4bbe3-1283">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1283">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4bbe3-1284">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1284">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4bbe3-1285">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1285">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="4bbe3-1286">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1286">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="4bbe3-1287">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1287">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="4bbe3-1288">Жеталлинтернесеадерсасинк ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1288">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="4bbe3-1289">Получает все заголовки Интернета для сообщения в виде строки.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1289">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="4bbe3-1290">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1290">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1291">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1291">Parameters</span></span>

|<span data-ttu-id="4bbe3-1292">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1292">Name</span></span>|<span data-ttu-id="4bbe3-1293">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1293">Type</span></span>|<span data-ttu-id="4bbe3-1294">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1294">Attributes</span></span>|<span data-ttu-id="4bbe3-1295">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1295">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4bbe3-1296">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1296">Object</span></span>|<span data-ttu-id="4bbe3-1297">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1297">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1298">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1298">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1299">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1299">Object</span></span>|<span data-ttu-id="4bbe3-1300">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1301">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1301">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1302">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1302">function</span></span>|<span data-ttu-id="4bbe3-1303">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1304">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1304">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="4bbe3-1305">В случае успешного выполнения данные заголовков Интернета предоставляются в свойстве asyncResult. Value в виде String.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1305">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="4bbe3-1306">Сведения о форматировании возвращаемого строкового значения приведены в [RFC 2183](https://tools.ietf.org/html/rfc2183) .</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1306">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="4bbe3-1307">Если происходит сбой вызова, свойство asyncResult. Error будет содержать код ошибки с причиной сбоя.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1307">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1308">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1308">Requirements</span></span>

|<span data-ttu-id="4bbe3-1309">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1309">Requirement</span></span>|<span data-ttu-id="4bbe3-1310">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1310">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1311">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1312">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1312">1.8</span></span>|
|[<span data-ttu-id="4bbe3-1313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1314">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1316">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1316">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1317">Returns:</span></span>

<span data-ttu-id="4bbe3-1318">Данные заголовков Интернета в виде строки, отформатированной в соответствии со [спецификацией RFC 2183](https://tools.ietf.org/html/rfc2183).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1318">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="4bbe3-1319">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1319">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1320">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1320">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="4bbe3-1321">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1321">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="4bbe3-1322">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1322">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="4bbe3-1323">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1323">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4bbe3-1324">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1324">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="4bbe3-1325">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1325">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4bbe3-1326">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1326">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1327">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1327">Parameters</span></span>

|<span data-ttu-id="4bbe3-1328">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1328">Name</span></span>|<span data-ttu-id="4bbe3-1329">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1329">Type</span></span>|<span data-ttu-id="4bbe3-1330">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1330">Attributes</span></span>|<span data-ttu-id="4bbe3-1331">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1331">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4bbe3-1332">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1332">String</span></span>||<span data-ttu-id="4bbe3-1333">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1333">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1334">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1334">Object</span></span>|<span data-ttu-id="4bbe3-1335">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1335">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1336">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1336">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1337">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1337">Object</span></span>|<span data-ttu-id="4bbe3-1338">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1338">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1339">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1339">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1340">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1340">function</span></span>|<span data-ttu-id="4bbe3-1341">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1342">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1343">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1343">Requirements</span></span>

|<span data-ttu-id="4bbe3-1344">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1344">Requirement</span></span>|<span data-ttu-id="4bbe3-1345">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1346">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1347">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1347">1.8</span></span>|
|[<span data-ttu-id="4bbe3-1348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1349">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1351">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1351">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1352">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1352">Returns:</span></span>

<span data-ttu-id="4bbe3-1353">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1353">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1354">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1354">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="4bbe3-1355">Жетаттачментсасинк ([параметры], [обратный вызов]) → массив. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-1355">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="4bbe3-1356">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1356">Gets the item's attachments as an array.</span></span> <span data-ttu-id="4bbe3-1357">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1357">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1358">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1358">Parameters</span></span>

|<span data-ttu-id="4bbe3-1359">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1359">Name</span></span>|<span data-ttu-id="4bbe3-1360">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1360">Type</span></span>|<span data-ttu-id="4bbe3-1361">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1361">Attributes</span></span>|<span data-ttu-id="4bbe3-1362">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1362">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4bbe3-1363">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1363">Object</span></span>|<span data-ttu-id="4bbe3-1364">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1364">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1365">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1365">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1366">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1366">Object</span></span>|<span data-ttu-id="4bbe3-1367">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1367">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1368">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1368">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1369">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1369">function</span></span>|<span data-ttu-id="4bbe3-1370">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1371">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1371">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1372">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1372">Requirements</span></span>

|<span data-ttu-id="4bbe3-1373">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1373">Requirement</span></span>|<span data-ttu-id="4bbe3-1374">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1374">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1375">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1376">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1376">1.8</span></span>|
|[<span data-ttu-id="4bbe3-1377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1378">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1380">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1380">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1381">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1381">Returns:</span></span>

<span data-ttu-id="4bbe3-1382">Тип: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4bbe3-1382">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1383">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1383">Example</span></span>

<span data-ttu-id="4bbe3-1384">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1384">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="4bbe3-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="4bbe3-1386">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1386">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1387">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1387">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1388">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1388">Requirements</span></span>

|<span data-ttu-id="4bbe3-1389">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1389">Requirement</span></span>|<span data-ttu-id="4bbe3-1390">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1390">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1391">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1392">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1392">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1393">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1393">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1394">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1395">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1395">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1396">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1396">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1397">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1397">Returns:</span></span>

<span data-ttu-id="4bbe3-1398">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1398">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1399">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1399">Example</span></span>

<span data-ttu-id="4bbe3-1400">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1400">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="4bbe3-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4bbe3-1402">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1402">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1403">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1403">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1404">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1404">Parameters</span></span>

|<span data-ttu-id="4bbe3-1405">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1405">Name</span></span>|<span data-ttu-id="4bbe3-1406">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1406">Type</span></span>|<span data-ttu-id="4bbe3-1407">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1407">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="4bbe3-1408">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1408">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="4bbe3-1409">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1409">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1410">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1410">Requirements</span></span>

|<span data-ttu-id="4bbe3-1411">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1411">Requirement</span></span>|<span data-ttu-id="4bbe3-1412">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1412">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1414">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1414">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1415">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1416">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1416">Restricted</span></span>|
|[<span data-ttu-id="4bbe3-1417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1418">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1418">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1419">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1419">Returns:</span></span>

<span data-ttu-id="4bbe3-1420">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1420">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4bbe3-1421">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1421">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4bbe3-1422">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1422">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4bbe3-1423">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1423">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="4bbe3-1424">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1424">Value of `entityType`</span></span>|<span data-ttu-id="4bbe3-1425">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1425">Type of objects in returned array</span></span>|<span data-ttu-id="4bbe3-1426">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1426">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="4bbe3-1427">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1427">String</span></span>|<span data-ttu-id="4bbe3-1428">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1428">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="4bbe3-1429">Contact</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1429">Contact</span></span>|<span data-ttu-id="4bbe3-1430">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1430">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="4bbe3-1431">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1431">String</span></span>|<span data-ttu-id="4bbe3-1432">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1432">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="4bbe3-1433">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1433">MeetingSuggestion</span></span>|<span data-ttu-id="4bbe3-1434">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1434">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="4bbe3-1435">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1435">PhoneNumber</span></span>|<span data-ttu-id="4bbe3-1436">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1436">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="4bbe3-1437">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1437">TaskSuggestion</span></span>|<span data-ttu-id="4bbe3-1438">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1438">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="4bbe3-1439">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1439">String</span></span>|<span data-ttu-id="4bbe3-1440">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1440">**Restricted**</span></span>|

<span data-ttu-id="4bbe3-1441">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4bbe3-1441">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1442">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1442">Example</span></span>

<span data-ttu-id="4bbe3-1443">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1443">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="4bbe3-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4bbe3-1445">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1445">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1446">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1446">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4bbe3-1447">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1447">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1448">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1448">Parameters</span></span>

|<span data-ttu-id="4bbe3-1449">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1449">Name</span></span>|<span data-ttu-id="4bbe3-1450">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1450">Type</span></span>|<span data-ttu-id="4bbe3-1451">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1451">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4bbe3-1452">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1452">String</span></span>|<span data-ttu-id="4bbe3-1453">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1453">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1454">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1454">Requirements</span></span>

|<span data-ttu-id="4bbe3-1455">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1455">Requirement</span></span>|<span data-ttu-id="4bbe3-1456">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1456">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1457">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1457">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1458">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1458">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1459">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1459">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1460">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1460">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1461">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1461">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1462">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1462">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1463">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1463">Returns:</span></span>

<span data-ttu-id="4bbe3-p174">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4bbe3-1466">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4bbe3-1466">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="4bbe3-1467">getInitializationContextAsync ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1467">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="4bbe3-1468">Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1468">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1469">Этот метод поддерживается только в Outlook 2016 или более поздней версии для Windows ("нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1469">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1470">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1470">Parameters</span></span>

|<span data-ttu-id="4bbe3-1471">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1471">Name</span></span>|<span data-ttu-id="4bbe3-1472">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1472">Type</span></span>|<span data-ttu-id="4bbe3-1473">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1473">Attributes</span></span>|<span data-ttu-id="4bbe3-1474">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1474">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4bbe3-1475">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1475">Object</span></span>|<span data-ttu-id="4bbe3-1476">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1476">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1477">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1477">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1478">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1478">Object</span></span>|<span data-ttu-id="4bbe3-1479">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1479">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1480">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1480">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1481">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1481">function</span></span>|<span data-ttu-id="4bbe3-1482">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1482">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1483">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1483">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bbe3-1484">При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1484">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="4bbe3-1485">Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1485">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1486">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1486">Requirements</span></span>

|<span data-ttu-id="4bbe3-1487">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1487">Requirement</span></span>|<span data-ttu-id="4bbe3-1488">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1489">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1490">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1490">Preview</span></span>|
|[<span data-ttu-id="4bbe3-1491">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1492">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1493">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1494">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1494">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-1495">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1495">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="4bbe3-1496">Жетитемидасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1496">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="4bbe3-1497">Асинхронно получает идентификатор сохраненного элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1497">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="4bbe3-1498">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1498">Compose mode only.</span></span>

<span data-ttu-id="4bbe3-1499">При вызове этот метод возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1499">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1500">Если надстройка вызывает `getItemIdAsync` элемент в режиме создания (например, чтобы получить доступ `itemId` к использованию с помощью EWS или REST API), имейте в виду, что если Outlook находится в режиме кэширования, может потребоваться некоторое время до синхронизации элемента с сервером.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1500">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="4bbe3-1501">Пока элемент не будет синхронизирован, он не `itemId` распознается и не будет использоваться, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1501">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1502">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1502">Parameters</span></span>

|<span data-ttu-id="4bbe3-1503">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1503">Name</span></span>|<span data-ttu-id="4bbe3-1504">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1504">Type</span></span>|<span data-ttu-id="4bbe3-1505">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1505">Attributes</span></span>|<span data-ttu-id="4bbe3-1506">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1506">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4bbe3-1507">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1507">Object</span></span>|<span data-ttu-id="4bbe3-1508">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1508">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1509">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1509">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1510">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1510">Object</span></span>|<span data-ttu-id="4bbe3-1511">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1511">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1512">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1512">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1513">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1513">function</span></span>||<span data-ttu-id="4bbe3-1514">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1514">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bbe3-1515">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1515">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bbe3-1516">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1516">Errors</span></span>

|<span data-ttu-id="4bbe3-1517">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1517">Error code</span></span>|<span data-ttu-id="4bbe3-1518">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1518">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="4bbe3-1519">Идентификатор невозможно извлечь, пока не будет сохранен элемент.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1519">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1520">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1520">Requirements</span></span>

|<span data-ttu-id="4bbe3-1521">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1521">Requirement</span></span>|<span data-ttu-id="4bbe3-1522">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1522">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1523">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1524">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1524">1.8</span></span>|
|[<span data-ttu-id="4bbe3-1525">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1526">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1527">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1528">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1528">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bbe3-1529">Примеры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1529">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4bbe3-1530">В следующем примере показана структура `result` параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1530">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="4bbe3-1531">`value` Свойство содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1531">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="4bbe3-1532">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1532">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4bbe3-1533">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1533">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1534">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1534">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4bbe3-p178">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4bbe3-1538">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1538">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4bbe3-1539">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1539">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4bbe3-p179">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1543">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1543">Requirements</span></span>

|<span data-ttu-id="4bbe3-1544">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1544">Requirement</span></span>|<span data-ttu-id="4bbe3-1545">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1546">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1547">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1547">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1548">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1549">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1550">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1551">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1551">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1552">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1552">Returns:</span></span>

<span data-ttu-id="4bbe3-p180">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4bbe3-1555">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1555">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4bbe3-1556">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1556">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4bbe3-1557">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1557">Example</span></span>

<span data-ttu-id="4bbe3-1558">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1558">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4bbe3-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4bbe3-1560">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1560">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1561">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1561">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4bbe3-1562">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1562">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4bbe3-p181">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1565">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1565">Parameters</span></span>

|<span data-ttu-id="4bbe3-1566">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1566">Name</span></span>|<span data-ttu-id="4bbe3-1567">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1567">Type</span></span>|<span data-ttu-id="4bbe3-1568">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1568">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4bbe3-1569">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1569">String</span></span>|<span data-ttu-id="4bbe3-1570">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1570">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1571">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1571">Requirements</span></span>

|<span data-ttu-id="4bbe3-1572">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1572">Requirement</span></span>|<span data-ttu-id="4bbe3-1573">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1573">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1574">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1575">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1575">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1576">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1577">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1578">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1579">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1579">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1580">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1580">Returns:</span></span>

<span data-ttu-id="4bbe3-1581">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1581">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="4bbe3-1582">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4bbe3-1582">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1583">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1583">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4bbe3-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4bbe3-1585">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1585">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4bbe3-p182">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает пустую строку для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p182">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1588">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1588">Parameters</span></span>

|<span data-ttu-id="4bbe3-1589">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1589">Name</span></span>|<span data-ttu-id="4bbe3-1590">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1590">Type</span></span>|<span data-ttu-id="4bbe3-1591">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1591">Attributes</span></span>|<span data-ttu-id="4bbe3-1592">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1592">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="4bbe3-1593">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1593">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4bbe3-p183">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1597">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1597">Object</span></span>|<span data-ttu-id="4bbe3-1598">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1599">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1599">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1600">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1600">Object</span></span>|<span data-ttu-id="4bbe3-1601">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1602">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1602">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1603">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1603">function</span></span>||<span data-ttu-id="4bbe3-1604">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1604">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bbe3-1605">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1605">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4bbe3-1606">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1606">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1607">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1607">Requirements</span></span>

|<span data-ttu-id="4bbe3-1608">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1608">Requirement</span></span>|<span data-ttu-id="4bbe3-1609">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1609">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1610">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1611">1.2</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1611">1.2</span></span>|
|[<span data-ttu-id="4bbe3-1612">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1613">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1614">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1615">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1615">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1616">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1616">Returns:</span></span>

<span data-ttu-id="4bbe3-1617">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1617">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="4bbe3-1618">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1618">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1619">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1619">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="4bbe3-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="4bbe3-1621">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1621">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="4bbe3-1622">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1622">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1623">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1623">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1624">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1624">Requirements</span></span>

|<span data-ttu-id="4bbe3-1625">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1625">Requirement</span></span>|<span data-ttu-id="4bbe3-1626">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1626">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1627">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1628">1.6</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1628">1.6</span></span>|
|[<span data-ttu-id="4bbe3-1629">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1630">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1631">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1632">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1632">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1633">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1633">Returns:</span></span>

<span data-ttu-id="4bbe3-1634">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1634">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1635">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1635">Example</span></span>

<span data-ttu-id="4bbe3-1636">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1636">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4bbe3-1637">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1637">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4bbe3-p186">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1640">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4bbe3-p187">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4bbe3-1644">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1644">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4bbe3-1645">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1645">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4bbe3-p188">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1649">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1649">Requirements</span></span>

|<span data-ttu-id="4bbe3-1650">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1650">Requirement</span></span>|<span data-ttu-id="4bbe3-1651">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1651">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1652">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1653">1.6</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1653">1.6</span></span>|
|[<span data-ttu-id="4bbe3-1654">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1655">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1656">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1657">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1657">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4bbe3-1658">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1658">Returns:</span></span>

<span data-ttu-id="4bbe3-p189">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4bbe3-1661">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1661">Example</span></span>

<span data-ttu-id="4bbe3-1662">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1662">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="4bbe3-1663">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1663">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="4bbe3-1664">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1664">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1665">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1665">Parameters</span></span>

|<span data-ttu-id="4bbe3-1666">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1666">Name</span></span>|<span data-ttu-id="4bbe3-1667">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1667">Type</span></span>|<span data-ttu-id="4bbe3-1668">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1668">Attributes</span></span>|<span data-ttu-id="4bbe3-1669">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1669">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4bbe3-1670">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1670">Object</span></span>|<span data-ttu-id="4bbe3-1671">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1671">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1672">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1672">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1673">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1673">Object</span></span>|<span data-ttu-id="4bbe3-1674">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1674">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1675">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1675">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1676">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1676">function</span></span>||<span data-ttu-id="4bbe3-1677">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bbe3-1678">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1678">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4bbe3-1679">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1679">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1680">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1680">Requirements</span></span>

|<span data-ttu-id="4bbe3-1681">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1681">Requirement</span></span>|<span data-ttu-id="4bbe3-1682">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1682">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1683">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1684">1.8</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1684">1.8</span></span>|
|[<span data-ttu-id="4bbe3-1685">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1686">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1687">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1688">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1688">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-1689">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1689">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4bbe3-1690">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1690">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4bbe3-1691">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1691">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4bbe3-p191">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1695">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1695">Parameters</span></span>

|<span data-ttu-id="4bbe3-1696">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1696">Name</span></span>|<span data-ttu-id="4bbe3-1697">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1697">Type</span></span>|<span data-ttu-id="4bbe3-1698">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1698">Attributes</span></span>|<span data-ttu-id="4bbe3-1699">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1699">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="4bbe3-1700">function</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1700">function</span></span>||<span data-ttu-id="4bbe3-1701">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bbe3-1702">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1702">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4bbe3-1703">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1703">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="4bbe3-1704">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1704">Object</span></span>|<span data-ttu-id="4bbe3-1705">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1705">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1706">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1706">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4bbe3-1707">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1707">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1708">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1708">Requirements</span></span>

|<span data-ttu-id="4bbe3-1709">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1709">Requirement</span></span>|<span data-ttu-id="4bbe3-1710">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1710">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1711">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1712">1.0</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1712">1.0</span></span>|
|[<span data-ttu-id="4bbe3-1713">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1714">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1714">ReadItem</span></span>|
|[<span data-ttu-id="4bbe3-1715">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1716">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1716">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-1717">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1717">Example</span></span>

<span data-ttu-id="4bbe3-p194">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4bbe3-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4bbe3-1722">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1722">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4bbe3-1723">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1723">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4bbe3-1724">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1724">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4bbe3-1725">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1725">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4bbe3-1726">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1726">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1727">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1727">Parameters</span></span>

|<span data-ttu-id="4bbe3-1728">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1728">Name</span></span>|<span data-ttu-id="4bbe3-1729">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1729">Type</span></span>|<span data-ttu-id="4bbe3-1730">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1730">Attributes</span></span>|<span data-ttu-id="4bbe3-1731">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1731">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4bbe3-1732">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1732">String</span></span>||<span data-ttu-id="4bbe3-1733">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1733">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1734">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1734">Object</span></span>|<span data-ttu-id="4bbe3-1735">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1735">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1736">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1736">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1737">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1737">Object</span></span>|<span data-ttu-id="4bbe3-1738">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1738">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1739">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1739">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1740">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1740">function</span></span>|<span data-ttu-id="4bbe3-1741">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1741">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1742">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1742">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4bbe3-1743">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1743">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4bbe3-1744">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1744">Errors</span></span>

|<span data-ttu-id="4bbe3-1745">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1745">Error code</span></span>|<span data-ttu-id="4bbe3-1746">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1746">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="4bbe3-1747">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1747">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1748">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1748">Requirements</span></span>

|<span data-ttu-id="4bbe3-1749">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1749">Requirement</span></span>|<span data-ttu-id="4bbe3-1750">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1750">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1751">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1752">1.1</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1752">1.1</span></span>|
|[<span data-ttu-id="4bbe3-1753">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1754">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1754">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-1755">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1756">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1756">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-1757">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1757">Example</span></span>

<span data-ttu-id="4bbe3-1758">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1758">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4bbe3-1759">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1759">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4bbe3-1760">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1760">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4bbe3-1761">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1761">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1762">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1762">Parameters</span></span>

| <span data-ttu-id="4bbe3-1763">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1763">Name</span></span> | <span data-ttu-id="4bbe3-1764">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1764">Type</span></span> | <span data-ttu-id="4bbe3-1765">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1765">Attributes</span></span> | <span data-ttu-id="4bbe3-1766">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1766">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4bbe3-1767">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1767">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4bbe3-1768">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1768">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4bbe3-1769">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1769">Object</span></span> | <span data-ttu-id="4bbe3-1770">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1770">&lt;optional&gt;</span></span> | <span data-ttu-id="4bbe3-1771">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1771">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4bbe3-1772">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1772">Object</span></span> | <span data-ttu-id="4bbe3-1773">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1773">&lt;optional&gt;</span></span> | <span data-ttu-id="4bbe3-1774">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1774">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4bbe3-1775">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1775">function</span></span>| <span data-ttu-id="4bbe3-1776">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1776">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1777">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1778">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1778">Requirements</span></span>

|<span data-ttu-id="4bbe3-1779">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1779">Requirement</span></span>| <span data-ttu-id="4bbe3-1780">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1780">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1781">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4bbe3-1782">1.7</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1782">1.7</span></span> |
|[<span data-ttu-id="4bbe3-1783">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4bbe3-1784">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1784">ReadItem</span></span> |
|[<span data-ttu-id="4bbe3-1785">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4bbe3-1786">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1786">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="4bbe3-1787">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1787">saveAsync([options], callback)</span></span>

<span data-ttu-id="4bbe3-1788">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1788">Asynchronously saves an item.</span></span>

<span data-ttu-id="4bbe3-1789">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1789">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="4bbe3-1790">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1790">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="4bbe3-1791">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1791">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1792">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1792">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4bbe3-1793">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1793">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4bbe3-p198">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4bbe3-1797">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1797">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4bbe3-1798">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1798">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="4bbe3-1799">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1799">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="4bbe3-1800">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1800">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="4bbe3-1801">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1801">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1802">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1802">Parameters</span></span>

|<span data-ttu-id="4bbe3-1803">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1803">Name</span></span>|<span data-ttu-id="4bbe3-1804">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1804">Type</span></span>|<span data-ttu-id="4bbe3-1805">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1805">Attributes</span></span>|<span data-ttu-id="4bbe3-1806">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1806">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4bbe3-1807">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1807">Object</span></span>|<span data-ttu-id="4bbe3-1808">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1808">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1809">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1809">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1810">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1810">Object</span></span>|<span data-ttu-id="4bbe3-1811">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1811">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1812">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1812">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1813">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1813">function</span></span>||<span data-ttu-id="4bbe3-1814">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1814">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4bbe3-1815">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1815">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1816">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1816">Requirements</span></span>

|<span data-ttu-id="4bbe3-1817">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1817">Requirement</span></span>|<span data-ttu-id="4bbe3-1818">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1818">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1819">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1819">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1820">1.3</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1820">1.3</span></span>|
|[<span data-ttu-id="4bbe3-1821">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1821">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1822">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1822">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-1823">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1823">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1824">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1824">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4bbe3-1825">Примеры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1825">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4bbe3-p200">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4bbe3-1828">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1828">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4bbe3-1829">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1829">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4bbe3-p201">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4bbe3-1833">Параметры</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1833">Parameters</span></span>

|<span data-ttu-id="4bbe3-1834">Имя</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1834">Name</span></span>|<span data-ttu-id="4bbe3-1835">Тип</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1835">Type</span></span>|<span data-ttu-id="4bbe3-1836">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1836">Attributes</span></span>|<span data-ttu-id="4bbe3-1837">Описание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1837">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="4bbe3-1838">String</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1838">String</span></span>||<span data-ttu-id="4bbe3-p202">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="4bbe3-1842">Object</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1842">Object</span></span>|<span data-ttu-id="4bbe3-1843">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1843">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1844">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1844">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4bbe3-1845">Объект</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1845">Object</span></span>|<span data-ttu-id="4bbe3-1846">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1846">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1847">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1847">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4bbe3-1848">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1848">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4bbe3-1849">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1849">&lt;optional&gt;</span></span>|<span data-ttu-id="4bbe3-1850">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1850">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="4bbe3-1851">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1851">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4bbe3-1852">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1852">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="4bbe3-1853">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1853">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4bbe3-1854">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1854">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="4bbe3-1855">функция</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1855">function</span></span>||<span data-ttu-id="4bbe3-1856">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bbe3-1857">Требования</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1857">Requirements</span></span>

|<span data-ttu-id="4bbe3-1858">Требование</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1858">Requirement</span></span>|<span data-ttu-id="4bbe3-1859">Значение</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1859">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bbe3-1860">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4bbe3-1861">1.2</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1861">1.2</span></span>|
|[<span data-ttu-id="4bbe3-1862">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4bbe3-1863">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1863">ReadWriteItem</span></span>|
|[<span data-ttu-id="4bbe3-1864">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4bbe3-1865">Создание</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1865">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4bbe3-1866">Пример</span><span class="sxs-lookup"><span data-stu-id="4bbe3-1866">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
