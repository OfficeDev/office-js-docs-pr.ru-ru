---
title: Общие сведения о разрешениях для надстроек Outlook
description: Надстройки Outlook указывают требуемый уровень разрешений в своем манифесте, который включает Restricted, ReadItem, ReadWriteItem, or ReadWriteMailbox.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 60b65416585b5215ed565a3689c1e7f398e001a5
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325328"
---
# <a name="understanding-outlook-add-in-permissions"></a><span data-ttu-id="405b5-103">Общие сведения о разрешениях для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="405b5-103">Understanding Outlook add-in permissions</span></span>

<span data-ttu-id="405b5-p101">Необходимый уровень разрешений для надстроек Outlook указывается в манифесте. Доступные уровни: **Restricted**, **ReadItem**, **ReadWriteItem** и **ReadWriteMailbox**. Эти уровни являются накопительными: **Restricted** — самый низкий уровень, каждый более высокий уровень включает разрешения более низких уровней. **ReadWriteMailbox** включает все поддерживаемые разрешения.</span><span class="sxs-lookup"><span data-stu-id="405b5-p101">Outlook add-ins specify the required permission level in their manifest. The available levels are **Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**. These levels of permissions are cumulative: **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.</span></span>

<span data-ttu-id="405b5-p102">Вы можете просмотреть разрешения, которые запрашивает почтовая надстройка, перед ее установкой из [AppSource](https://appsource.microsoft.com). Вы также можете просмотреть требуемые разрешения установленных надстроек в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="405b5-p102">You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.</span></span>

## <a name="restricted-permission"></a><span data-ttu-id="405b5-110">Разрешение Restricted</span><span class="sxs-lookup"><span data-stu-id="405b5-110">Restricted permission</span></span>

<span data-ttu-id="405b5-p103">
  \*\*Restricted\*\* — самый простой уровень разрешений. Укажите \*\*Restricted\*\* в элементе [Permissions](../reference/manifest/permissions.md) манифеста, чтобы запросить это разрешение. Outlook назначает это разрешение почтовой надстройке по умолчанию, если надстройка не запрашивает особого разрешения в манифесте.</span><span class="sxs-lookup"><span data-stu-id="405b5-p103">The **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](../reference/manifest/permissions.md) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.</span></span>

### <a name="can-do"></a><span data-ttu-id="405b5-114">Разрешено</span><span class="sxs-lookup"><span data-stu-id="405b5-114">Can do</span></span>

- <span data-ttu-id="405b5-115">[Получать только определенные сущности](match-strings-in-an-item-as-well-known-entities.md) (номер телефона, адрес, URL-адрес) из темы или текста элемента.</span><span class="sxs-lookup"><span data-stu-id="405b5-115">[Get only specific entities](match-strings-in-an-item-as-well-known-entities.md) (phone number, address, URL) from the item's subject or body.</span></span>

- <span data-ttu-id="405b5-116">Указывать [правило активации ItemIs](activation-rules.md#itemis-rule), требующее, чтобы текущий элемент в форме чтения или создания принадлежал определенному типу, или правило [ItemHasKnownEntity](match-strings-in-an-item-as-well-known-entities.md), соответствующее малому поднабору поддерживаемых известных сущностей (номер телефона, адрес, URL-адрес) в выбранном элементе.</span><span class="sxs-lookup"><span data-stu-id="405b5-116">Specify an [ItemIs activation rule](activation-rules.md#itemis-rule) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.</span></span>

- <span data-ttu-id="405b5-117">Получать доступ к свойствам и методам, которые **не** относятся к определенной информации о пользователе или элементе (список элементов, которые относятся к такой информации, см. в следующем разделе).</span><span class="sxs-lookup"><span data-stu-id="405b5-117">Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).</span></span>

### <a name="cant-do"></a><span data-ttu-id="405b5-118">Не разрешено</span><span class="sxs-lookup"><span data-stu-id="405b5-118">Can't do</span></span>

- <span data-ttu-id="405b5-119">Используйте правило [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) для контакта, адрес электронной почты, предложение о собрании или сущность предложения по задаче.</span><span class="sxs-lookup"><span data-stu-id="405b5-119">Use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the contact, email address, meeting suggestion, or task suggestion entity.</span></span>

- <span data-ttu-id="405b5-120">Использовать правило [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) или [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule).</span><span class="sxs-lookup"><span data-stu-id="405b5-120">Use the [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule.</span></span>

- <span data-ttu-id="405b5-p104">Получать доступ к элементам в приведенном ниже списке, которые относятся к информации о пользователе или элементе. При попытке получить доступ к элементам в этом списке будут возвращены значение **null** и сообщение о том, что требуются повышенные привилегии.</span><span class="sxs-lookup"><span data-stu-id="405b5-p104">Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.</span></span>

    - [<span data-ttu-id="405b5-123">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-123">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-124">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-124">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-125">item.attachments</span><span class="sxs-lookup"><span data-stu-id="405b5-125">item.attachments</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-126">item.bcc</span><span class="sxs-lookup"><span data-stu-id="405b5-126">item.bcc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-127">item.body</span><span class="sxs-lookup"><span data-stu-id="405b5-127">item.body</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-128">item.cc</span><span class="sxs-lookup"><span data-stu-id="405b5-128">item.cc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-129">item.from</span><span class="sxs-lookup"><span data-stu-id="405b5-129">item.from</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-130">item.getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="405b5-130">item.getRegExMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-131">item.getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="405b5-131">item.getRegExMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-132">item.optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="405b5-132">item.optionalAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-133">item.organizer</span><span class="sxs-lookup"><span data-stu-id="405b5-133">item.organizer</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-134">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-134">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-135">item.requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="405b5-135">item.requiredAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-136">item.sender</span><span class="sxs-lookup"><span data-stu-id="405b5-136">item.sender</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-137">item.to</span><span class="sxs-lookup"><span data-stu-id="405b5-137">item.to</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="405b5-138">mailbox.getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-138">mailbox.getCallbackTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="405b5-139">mailbox.getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-139">mailbox.getUserIdentityTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="405b5-140">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-140">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="405b5-141">mailbox.userProfile</span><span class="sxs-lookup"><span data-stu-id="405b5-141">mailbox.userProfile</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - <span data-ttu-id="405b5-142">[Body](/javascript/api/outlook/office.body) и все дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="405b5-142">[Body](/javascript/api/outlook/office.body) and all its child members</span></span>
    - <span data-ttu-id="405b5-143">[Location](/javascript/api/outlook/office.location) и все дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="405b5-143">[Location](/javascript/api/outlook/office.location) and all its child members</span></span>
    - <span data-ttu-id="405b5-144">[Recipients](/javascript/api/outlook/office.recipients) и все дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="405b5-144">[Recipients](/javascript/api/outlook/office.recipients) and all its child members</span></span>
    - <span data-ttu-id="405b5-145">[Subject](/javascript/api/outlook/office.subject) и все дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="405b5-145">[Subject](/javascript/api/outlook/office.subject) and all its child members</span></span>
    - <span data-ttu-id="405b5-146">[Time](/javascript/api/outlook/office.time) и все дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="405b5-146">[Time](/javascript/api/outlook/office.time) and all its child members</span></span>

## <a name="readitem-permission"></a><span data-ttu-id="405b5-147">Разрешение ReadItem</span><span class="sxs-lookup"><span data-stu-id="405b5-147">ReadItem permission</span></span>

<span data-ttu-id="405b5-p105">**ReadItem** — следующий уровень в модели разрешений. Укажите **ReadItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение.</span><span class="sxs-lookup"><span data-stu-id="405b5-p105">The **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="405b5-150">Разрешено</span><span class="sxs-lookup"><span data-stu-id="405b5-150">Can do</span></span>

- <span data-ttu-id="405b5-151">[Считывать все свойства](item-data.md) текущего элемента в чтении или [Создавать форму](get-and-set-item-data-in-a-compose-form.md), например [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) в форме чтения и [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) в форме создания.</span><span class="sxs-lookup"><span data-stu-id="405b5-151">[Read all the properties](item-data.md) of the current item in a read or [compose form](get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) in a compose form.</span></span>

- <span data-ttu-id="405b5-152">[Получать маркер обратного вызова для получения вложений](get-attachments-of-an-outlook-item.md) или всего элемента с помощью веб-служб Exchange или [REST API Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="405b5-152">[Get a callback token to get item attachments](get-attachments-of-an-outlook-item.md) or the full item with Exchange Web Services (EWS) or [Outlook REST APIs](use-rest-api.md).</span></span>

- <span data-ttu-id="405b5-153">[Записывать пользовательские свойства](/javascript/api/outlook/office.CustomProperties), установленные надстройкой для соответствующего элемента.</span><span class="sxs-lookup"><span data-stu-id="405b5-153">[Write custom properties](/javascript/api/outlook/office.CustomProperties) set by the add-in on that item.</span></span>

- <span data-ttu-id="405b5-154">[Получать все существующие известные сущности](match-strings-in-an-item-as-well-known-entities.md) (а не только группу) из темы или текста элемента.</span><span class="sxs-lookup"><span data-stu-id="405b5-154">[Get all existing well-known entities](match-strings-in-an-item-as-well-known-entities.md), not just a subset, from the item's subject or body.</span></span>

- <span data-ttu-id="405b5-p106">Использовать все [известные сущности](activation-rules.md#itemhasknownentity-rule) в правилах [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) или [регулярные выражения](activation-rules.md#itemhasregularexpressionmatch-rule) в правилах [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule). Следующий пример, использующий схему версии 1.1, активирует надстройку, если обнаруживается одна или несколько известных сущностей в теме или теле выбранного сообщения:</span><span class="sxs-lookup"><span data-stu-id="405b5-p106">Use all the [well-known entities](activation-rules.md#itemhasknownentity-rule) in [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rules, or [regular expressions](activation-rules.md#itemhasregularexpressionmatch-rule) in [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a><span data-ttu-id="405b5-158">Не разрешено</span><span class="sxs-lookup"><span data-stu-id="405b5-158">Can't do</span></span>

- <span data-ttu-id="405b5-159">Использовать токен, предоставляемый методом **mailbox.getCallbackTokenAsync**, для следующего:</span><span class="sxs-lookup"><span data-stu-id="405b5-159">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="405b5-160">обновление или удаление текущего элемента с помощью REST API для Outlook и получение доступа к другим элементам в почтовом ящике пользователя;</span><span class="sxs-lookup"><span data-stu-id="405b5-160">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="405b5-161">получение текущего элемента события календаря с помощью REST API для Outlook.</span><span class="sxs-lookup"><span data-stu-id="405b5-161">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="405b5-162">Использовать один из следующих API:</span><span class="sxs-lookup"><span data-stu-id="405b5-162">Use any of the following APIs:</span></span>
    - [<span data-ttu-id="405b5-163">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-163">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="405b5-164">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-164">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-165">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-165">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-166">item.bcc.addAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-166">item.bcc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-167">item.bcc.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-167">item.bcc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-168">item.body.prependAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-168">item.body.prependAsync</span></span>](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [<span data-ttu-id="405b5-169">item.body.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-169">item.body.setAsync</span></span>](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [<span data-ttu-id="405b5-170">item.body.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-170">item.body.setSelectedDataAsync</span></span>](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [<span data-ttu-id="405b5-171">item.cc.addAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-171">item.cc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-172">item.cc.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-172">item.cc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-173">item.end.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-173">item.end.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="405b5-174">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-174">item.location.setAsync</span></span>](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [<span data-ttu-id="405b5-175">item.optionalAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-175">item.optionalAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-176">item.optionalAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-176">item.optionalAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-177">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-177">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="405b5-178">item.requiredAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-178">item.requiredAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-179">item.requiredAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-179">item.requiredAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-180">item.start.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-180">item.start.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="405b5-181">item.subject.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-181">item.subject.setAsync</span></span>](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [<span data-ttu-id="405b5-182">item.to.addAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-182">item.to.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="405b5-183">item.to.setAsync</span><span class="sxs-lookup"><span data-stu-id="405b5-183">item.to.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a><span data-ttu-id="405b5-184">Разрешение ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="405b5-184">ReadWriteItem permission</span></span>

<span data-ttu-id="405b5-p107">Укажите элемент **ReadWriteItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение. Почтовые надстройки, активированные в формах создания, которые используют методы записи (**Message.to.addAsync** или **Message.to.setAsync**), должны использовать по крайней мере этот уровень разрешений.</span><span class="sxs-lookup"><span data-stu-id="405b5-p107">Specify **ReadWriteItem** in the **Permissions** element in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="405b5-187">Разрешено</span><span class="sxs-lookup"><span data-stu-id="405b5-187">Can do</span></span>

- <span data-ttu-id="405b5-188">[Считывать и записывать все свойства на уровне элемента](item-data.md) для элемента, который просматривается или создается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="405b5-188">[Read and write all item-level properties](item-data.md) of the item that is being viewed or composed in Outlook.</span></span>

- <span data-ttu-id="405b5-189">[Добавлять или удалять вложения](add-and-remove-attachments-to-an-item-in-a-compose-form.md) для такого элемента.</span><span class="sxs-lookup"><span data-stu-id="405b5-189">[Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.</span></span>

- <span data-ttu-id="405b5-190">Используйте все остальные элементы API JavaScript для Office, которые относятся к почтовым надстройкам, за исключением **Mailbox. makeEWSRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="405b5-190">Use all other members of the Office JavaScript API that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.</span></span>

### <a name="cant-do"></a><span data-ttu-id="405b5-191">Не разрешено</span><span class="sxs-lookup"><span data-stu-id="405b5-191">Can't do</span></span>

- <span data-ttu-id="405b5-192">Использовать токен, предоставляемый методом **mailbox.getCallbackTokenAsync**, для следующего:</span><span class="sxs-lookup"><span data-stu-id="405b5-192">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="405b5-193">обновление или удаление текущего элемента с помощью REST API для Outlook и получение доступа к другим элементам в почтовом ящике пользователя;</span><span class="sxs-lookup"><span data-stu-id="405b5-193">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="405b5-194">получение текущего элемента события календаря с помощью REST API для Outlook.</span><span class="sxs-lookup"><span data-stu-id="405b5-194">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="405b5-195">Использовать **mailbox.makeEWSRequestAsync**.</span><span class="sxs-lookup"><span data-stu-id="405b5-195">Use **mailbox.makeEWSRequestAsync**.</span></span>

## <a name="readwritemailbox-permission"></a><span data-ttu-id="405b5-196">Разрешение ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="405b5-196">ReadWriteMailbox permission</span></span>

<span data-ttu-id="405b5-p108">**ReadWriteMailbox** — самый высокий уровень разрешений. Укажите **ReadWriteMailbox** в элементе **Permissions** манифеста, чтобы запросить это разрешение.</span><span class="sxs-lookup"><span data-stu-id="405b5-p108">The **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.</span></span>

<span data-ttu-id="405b5-199">В дополнение к тому, что поддерживает разрешение **ReadWriteItem**, токен, предоставляемый элементом **mailbox.getCallbackTokenAsync**, позволяет использовать операции веб-служб Exchange или REST API Outlook для выполнения следующих действий:</span><span class="sxs-lookup"><span data-stu-id="405b5-199">In addition to what the **ReadWriteItem** permission supports, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:</span></span>

- <span data-ttu-id="405b5-200">Чтение и запись всех свойств любого элемента в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="405b5-200">Read and write all properties of any item in the user's mailbox.</span></span>
- <span data-ttu-id="405b5-201">Создание, чтение и запись в любую папку или элемент в таком почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="405b5-201">Create, read, and write to any folder or item in that mailbox.</span></span>
- <span data-ttu-id="405b5-202">Отправка элемента из такого почтового ящика</span><span class="sxs-lookup"><span data-stu-id="405b5-202">Send an item from that mailbox</span></span>

<span data-ttu-id="405b5-203">С помощью **mailbox.makeEWSRequestAsync** вы можете использовать следующие операции EWS:</span><span class="sxs-lookup"><span data-stu-id="405b5-203">Through **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:</span></span>

- [<span data-ttu-id="405b5-204">CopyItem</span><span class="sxs-lookup"><span data-stu-id="405b5-204">CopyItem</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)
- [<span data-ttu-id="405b5-205">CreateFolder</span><span class="sxs-lookup"><span data-stu-id="405b5-205">CreateFolder</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)
- [<span data-ttu-id="405b5-206">CreateItem</span><span class="sxs-lookup"><span data-stu-id="405b5-206">CreateItem</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)
- [<span data-ttu-id="405b5-207">FindConversation</span><span class="sxs-lookup"><span data-stu-id="405b5-207">FindConversation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)
- [<span data-ttu-id="405b5-208">FindFolder</span><span class="sxs-lookup"><span data-stu-id="405b5-208">FindFolder</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)
- [<span data-ttu-id="405b5-209">FindItem</span><span class="sxs-lookup"><span data-stu-id="405b5-209">FindItem</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)
- [<span data-ttu-id="405b5-210">GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="405b5-210">GetConversationItems</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [<span data-ttu-id="405b5-211">GetFolder</span><span class="sxs-lookup"><span data-stu-id="405b5-211">GetFolder</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)
- [<span data-ttu-id="405b5-212">GetItem</span><span class="sxs-lookup"><span data-stu-id="405b5-212">GetItem</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)
- [<span data-ttu-id="405b5-213">MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="405b5-213">MarkAsJunk</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [<span data-ttu-id="405b5-214">MoveItem</span><span class="sxs-lookup"><span data-stu-id="405b5-214">MoveItem</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)
- [<span data-ttu-id="405b5-215">SendItem</span><span class="sxs-lookup"><span data-stu-id="405b5-215">SendItem</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)
- [<span data-ttu-id="405b5-216">UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="405b5-216">UpdateFolder</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [<span data-ttu-id="405b5-217">UpdateItem</span><span class="sxs-lookup"><span data-stu-id="405b5-217">UpdateItem</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)

<span data-ttu-id="405b5-218">Попытка использования неподдерживаемой операции приведет к возврату ошибки.</span><span class="sxs-lookup"><span data-stu-id="405b5-218">Attempting to use an unsupported operation will result in an error response.</span></span>

## <a name="see-also"></a><span data-ttu-id="405b5-219">См. также</span><span class="sxs-lookup"><span data-stu-id="405b5-219">See also</span></span>

- [<span data-ttu-id="405b5-220">Конфиденциальность, разрешения и безопасность для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="405b5-220">Privacy, permissions, and security for Outlook add-ins</span></span>](../develop/privacy-and-security.md)
- [<span data-ttu-id="405b5-221">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="405b5-221">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
