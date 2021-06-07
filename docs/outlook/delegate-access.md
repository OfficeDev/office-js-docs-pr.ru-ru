---
title: Включить сценарии делегирования доступа в Outlook надстройки
description: Кратко описывает делегатский доступ и рассказывает о настройке поддержки надстройки.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 256c37087b10eaf9c8025e19a4990852f9550458
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783493"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="46cf4-103">Включить сценарии делегирования доступа в Outlook надстройки</span><span class="sxs-lookup"><span data-stu-id="46cf4-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="46cf4-104">Владелец почтового ящика может использовать функцию доступа к делегатам, чтобы позволить другому человеку управлять [своей почтой и календарем.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)</span><span class="sxs-lookup"><span data-stu-id="46cf4-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="46cf4-105">В этой статье указывается, какие разрешения делегировать Office API JavaScript поддерживает, и описывается, как включить сценарии делегирования доступа Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="46cf4-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="46cf4-106">В настоящее время доступ к делегированию Outlook на Android и iOS.</span><span class="sxs-lookup"><span data-stu-id="46cf4-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="46cf4-107">Кроме того, эта функция [](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) в настоящее время недоступна для групповых общих почтовых ящиков Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="46cf4-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="46cf4-108">Эта функция может быть доступна в будущем.</span><span class="sxs-lookup"><span data-stu-id="46cf4-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="46cf4-109">Поддержка этой функции была представлена в наборе требований 1.8.</span><span class="sxs-lookup"><span data-stu-id="46cf4-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="46cf4-110">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="46cf4-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="46cf4-111">Поддерживаемые разрешения для доступа к делегированию</span><span class="sxs-lookup"><span data-stu-id="46cf4-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="46cf4-112">В следующей таблице описываются разрешения делегатов, которые Office API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="46cf4-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="46cf4-113">Разрешение</span><span class="sxs-lookup"><span data-stu-id="46cf4-113">Permission</span></span>|<span data-ttu-id="46cf4-114">Значение</span><span class="sxs-lookup"><span data-stu-id="46cf4-114">Value</span></span>|<span data-ttu-id="46cf4-115">Описание</span><span class="sxs-lookup"><span data-stu-id="46cf4-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="46cf4-116">Чтение</span><span class="sxs-lookup"><span data-stu-id="46cf4-116">Read</span></span>|<span data-ttu-id="46cf4-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="46cf4-117">1 (000001)</span></span>|<span data-ttu-id="46cf4-118">Может читать элементы.</span><span class="sxs-lookup"><span data-stu-id="46cf4-118">Can read items.</span></span>|
|<span data-ttu-id="46cf4-119">Запись</span><span class="sxs-lookup"><span data-stu-id="46cf4-119">Write</span></span>|<span data-ttu-id="46cf4-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="46cf4-120">2 (000010)</span></span>|<span data-ttu-id="46cf4-121">Можно создавать элементы.</span><span class="sxs-lookup"><span data-stu-id="46cf4-121">Can create items.</span></span>|
|<span data-ttu-id="46cf4-122">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="46cf4-122">DeleteOwn</span></span>|<span data-ttu-id="46cf4-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="46cf4-123">4 (000100)</span></span>|<span data-ttu-id="46cf4-124">Можно удалить только созданные элементы.</span><span class="sxs-lookup"><span data-stu-id="46cf4-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="46cf4-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="46cf4-125">DeleteAll</span></span>|<span data-ttu-id="46cf4-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="46cf4-126">8 (001000)</span></span>|<span data-ttu-id="46cf4-127">Может удалять любые элементы.</span><span class="sxs-lookup"><span data-stu-id="46cf4-127">Can delete any items.</span></span>|
|<span data-ttu-id="46cf4-128">EditOwn</span><span class="sxs-lookup"><span data-stu-id="46cf4-128">EditOwn</span></span>|<span data-ttu-id="46cf4-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="46cf4-129">16 (010000)</span></span>|<span data-ttu-id="46cf4-130">Можно редактировать только созданные элементы.</span><span class="sxs-lookup"><span data-stu-id="46cf4-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="46cf4-131">EditAll</span><span class="sxs-lookup"><span data-stu-id="46cf4-131">EditAll</span></span>|<span data-ttu-id="46cf4-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="46cf4-132">32 (100000)</span></span>|<span data-ttu-id="46cf4-133">Может изменять любые элементы.</span><span class="sxs-lookup"><span data-stu-id="46cf4-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="46cf4-134">В настоящее время API поддерживает получение существующих разрешений делегирования, но не установку разрешений делегирования.</span><span class="sxs-lookup"><span data-stu-id="46cf4-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="46cf4-135">Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битмаски для указать разрешения делегата.</span><span class="sxs-lookup"><span data-stu-id="46cf4-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="46cf4-136">Каждая позиция в битмаске представляет определенное разрешение, и если оно заданной, то делегат `1` имеет соответствующее разрешение.</span><span class="sxs-lookup"><span data-stu-id="46cf4-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="46cf4-137">Например, если второй бит справа , то у делегата `1` есть разрешение **на записи.**</span><span class="sxs-lookup"><span data-stu-id="46cf4-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="46cf4-138">Пример проверки определенного разрешения в разделе [Выполнение](#perform-an-operation-as-delegate) операции в качестве делегата см. в этой статье.</span><span class="sxs-lookup"><span data-stu-id="46cf4-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="46cf4-139">Синхронизация между клиентами почтовых ящиков</span><span class="sxs-lookup"><span data-stu-id="46cf4-139">Sync across mailbox clients</span></span>

<span data-ttu-id="46cf4-140">Обновления делегата в почтовом ящике владельца обычно синхронизируются между почтовыми ящиками немедленно.</span><span class="sxs-lookup"><span data-stu-id="46cf4-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="46cf4-141">Однако если операции REST или Exchange Web Services (EWS) использовались для набора расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Мы рекомендуем вместо этого использовать [объект CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки.</span><span class="sxs-lookup"><span data-stu-id="46cf4-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="46cf4-142">Дополнительные дополнительные [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи см. в разделе настраиваемые свойства в статье "Получить и установить метаданные в Outlook надстройки".</span><span class="sxs-lookup"><span data-stu-id="46cf4-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="46cf4-143">В сценарии делегирования нельзя использовать EWS с маркерами, которые в настоящее время office.js API.</span><span class="sxs-lookup"><span data-stu-id="46cf4-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="46cf4-144">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="46cf4-144">Configure the manifest</span></span>

<span data-ttu-id="46cf4-145">Чтобы включить сценарии делегирования доступа в надстройку, необходимо настроить элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) в манифесте под `true` родительским элементом. `DesktopFormFactor`</span><span class="sxs-lookup"><span data-stu-id="46cf4-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="46cf4-146">В настоящее время другие форм-факторы не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="46cf4-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="46cf4-147">Чтобы поддерживать вызовы REST от делегата, установите узел [Разрешений](../reference/manifest/permissions.md) в `ReadWriteMailbox` манифесте.</span><span class="sxs-lookup"><span data-stu-id="46cf4-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="46cf4-148">В следующем примере показан элемент, установленный `SupportsSharedFolders` `true` в разделе манифеста.</span><span class="sxs-lookup"><span data-stu-id="46cf4-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="46cf4-149">Выполнение операции в качестве делегата</span><span class="sxs-lookup"><span data-stu-id="46cf4-149">Perform an operation as delegate</span></span>

<span data-ttu-id="46cf4-150">Общие свойства элемента можно получить в режиме Compose или Read, позвонив по методу [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="46cf4-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="46cf4-151">Это возвращает объект [SharedProperties,](/javascript/api/outlook/office.sharedproperties) который в настоящее время предоставляет разрешения делегата, адрес электронной почты владельца, базовый URL-адрес API REST и целевой почтовый ящик.</span><span class="sxs-lookup"><span data-stu-id="46cf4-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="46cf4-152">В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата разрешение **на** запись, и сделать вызов REST.</span><span class="sxs-lookup"><span data-stu-id="46cf4-152">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="46cf4-153">В качестве делегата можно использовать REST для получения содержимого сообщения Outlook, прикрепленного к элементу Outlook [или групповой публикации.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)</span><span class="sxs-lookup"><span data-stu-id="46cf4-153">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="46cf4-154">Обработка вызовов REST для общих и не общих элементов</span><span class="sxs-lookup"><span data-stu-id="46cf4-154">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="46cf4-155">Если вы хотите вызвать операцию REST для элемента, является ли этот элемент общим, вы можете использовать API, чтобы определить, является ли элемент `getSharedPropertiesAsync` общим.</span><span class="sxs-lookup"><span data-stu-id="46cf4-155">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="46cf4-156">После этого можно создать URL-адрес REST для операции с помощью соответствующего объекта.</span><span class="sxs-lookup"><span data-stu-id="46cf4-156">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="46cf4-157">Ограничения</span><span class="sxs-lookup"><span data-stu-id="46cf4-157">Limitations</span></span>

<span data-ttu-id="46cf4-158">В зависимости от сценариев надстройки существует несколько ограничений, которые следует учитывать при работе с ситуациями делегатов.</span><span class="sxs-lookup"><span data-stu-id="46cf4-158">Depending on your add-in's scenarios, there are a couple of limitations for you to consider when handling delegate situations.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="46cf4-159">REST и EWS</span><span class="sxs-lookup"><span data-stu-id="46cf4-159">REST and EWS</span></span>

<span data-ttu-id="46cf4-160">Ваша надстройка может использовать REST, но не EWS, и необходимо установить разрешение надстройки, чтобы включить доступ REST к почтовому `ReadWriteMailbox` ящику владельца.</span><span class="sxs-lookup"><span data-stu-id="46cf4-160">Your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="46cf4-161">Режим композитации сообщений</span><span class="sxs-lookup"><span data-stu-id="46cf4-161">Message Compose mode</span></span>

<span data-ttu-id="46cf4-162">В режиме композитации сообщений [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) не поддерживается Outlook в Интернете или Windows, если не выполнены следующие условия.</span><span class="sxs-lookup"><span data-stu-id="46cf4-162">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or Windows unless the following conditions are met.</span></span>

1. <span data-ttu-id="46cf4-163">Владелец делит с делегатом по крайней мере одну папку почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="46cf4-163">The owner shares at least one mailbox folder with the delegate.</span></span>
1. <span data-ttu-id="46cf4-164">Делегат проектирует сообщение в общей папке.</span><span class="sxs-lookup"><span data-stu-id="46cf4-164">The delegate drafts a message in the shared folder.</span></span>

    <span data-ttu-id="46cf4-165">Примеры:</span><span class="sxs-lookup"><span data-stu-id="46cf4-165">Examples:</span></span>

    - <span data-ttu-id="46cf4-166">Делегат отвечает на сообщения электронной почты в общей папке или переададирует их.</span><span class="sxs-lookup"><span data-stu-id="46cf4-166">The delegate replies to or forwards an email in the shared folder.</span></span>
    - <span data-ttu-id="46cf4-167">Делегат сохраняет черновик сообщения, а затем перемещает его из собственной папки **Drafts** в общую папку.</span><span class="sxs-lookup"><span data-stu-id="46cf4-167">The delegate saves a draft message then moves it from their own **Drafts** folder to the shared folder.</span></span> <span data-ttu-id="46cf4-168">Делегат открывает черновик из общей папки, а затем продолжает сочинять.</span><span class="sxs-lookup"><span data-stu-id="46cf4-168">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="46cf4-169">После того как сообщение отправлено, оно обычно находится в папке **отправленных** элементов делегата.</span><span class="sxs-lookup"><span data-stu-id="46cf4-169">After the message has been sent, it's usually found in the delegate's **Sent Items** folder.</span></span>

## <a name="see-also"></a><span data-ttu-id="46cf4-170">См. также</span><span class="sxs-lookup"><span data-stu-id="46cf4-170">See also</span></span>

- [<span data-ttu-id="46cf4-171">Разрешить другим пользователям управлять почтой и календарем</span><span class="sxs-lookup"><span data-stu-id="46cf4-171">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="46cf4-172">Общий доступ к календарю в Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="46cf4-172">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="46cf4-173">Как заказать элементы манифеста</span><span class="sxs-lookup"><span data-stu-id="46cf4-173">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="46cf4-174">[Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="46cf4-174">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="46cf4-175">Операторы bitwise JavaScript</span><span class="sxs-lookup"><span data-stu-id="46cf4-175">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)