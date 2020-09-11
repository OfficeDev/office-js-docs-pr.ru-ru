---
title: Включение сценариев делегирования доступа в надстройке Outlook
description: В кратко описывается доступ представителя и описывается настройка поддержки надстройки.
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 68b912d35f68cbf1177dd0b809994840092330a9
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430984"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="1c780-103">Включение сценариев делегирования доступа в надстройке Outlook</span><span class="sxs-lookup"><span data-stu-id="1c780-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="1c780-104">Владелец почтового ящика может использовать функцию делегированного доступа, чтобы [Разрешить другому пользователю управлять своей почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="1c780-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="1c780-105">В этой статье указывается, какие разрешения представителей поддерживает API JavaScript для Office, а также описывается включение сценариев делегированного доступа в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="1c780-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1c780-106">Доступ к представителю в настоящее время недоступен в Outlook на Android и iOS.</span><span class="sxs-lookup"><span data-stu-id="1c780-106">Delegate access is not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="1c780-107">Кроме того, эта функция в настоящее время недоступна для [групп общих почтовых ящиков](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="1c780-107">Also, this feature is not currently available with [group shared mailboxes](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) in Outlook on the web.</span></span> <span data-ttu-id="1c780-108">Эта функция может быть доступна в будущем.</span><span class="sxs-lookup"><span data-stu-id="1c780-108">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="1c780-109">Поддержка этой функции появилась в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="1c780-109">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="1c780-110">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="1c780-110">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="1c780-111">Поддерживаемые разрешения для делегированного доступа</span><span class="sxs-lookup"><span data-stu-id="1c780-111">Supported permissions for delegate access</span></span>

<span data-ttu-id="1c780-112">В следующей таблице описаны разрешения представителей, поддерживаемые API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="1c780-112">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="1c780-113">Разрешение</span><span class="sxs-lookup"><span data-stu-id="1c780-113">Permission</span></span>|<span data-ttu-id="1c780-114">Значение</span><span class="sxs-lookup"><span data-stu-id="1c780-114">Value</span></span>|<span data-ttu-id="1c780-115">Описание</span><span class="sxs-lookup"><span data-stu-id="1c780-115">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="1c780-116">Чтение</span><span class="sxs-lookup"><span data-stu-id="1c780-116">Read</span></span>|<span data-ttu-id="1c780-117">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="1c780-117">1 (000001)</span></span>|<span data-ttu-id="1c780-118">Возможность чтения элементов.</span><span class="sxs-lookup"><span data-stu-id="1c780-118">Can read items.</span></span>|
|<span data-ttu-id="1c780-119">Запись</span><span class="sxs-lookup"><span data-stu-id="1c780-119">Write</span></span>|<span data-ttu-id="1c780-120">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="1c780-120">2 (000010)</span></span>|<span data-ttu-id="1c780-121">Может создавать элементы.</span><span class="sxs-lookup"><span data-stu-id="1c780-121">Can create items.</span></span>|
|<span data-ttu-id="1c780-122">делетеовн</span><span class="sxs-lookup"><span data-stu-id="1c780-122">DeleteOwn</span></span>|<span data-ttu-id="1c780-123">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="1c780-123">4 (000100)</span></span>|<span data-ttu-id="1c780-124">Можно удалять только созданные ими элементы.</span><span class="sxs-lookup"><span data-stu-id="1c780-124">Can delete only the items they created.</span></span>|
|<span data-ttu-id="1c780-125">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="1c780-125">DeleteAll</span></span>|<span data-ttu-id="1c780-126">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="1c780-126">8 (001000)</span></span>|<span data-ttu-id="1c780-127">Может удалять все элементы.</span><span class="sxs-lookup"><span data-stu-id="1c780-127">Can delete any items.</span></span>|
|<span data-ttu-id="1c780-128">едитовн</span><span class="sxs-lookup"><span data-stu-id="1c780-128">EditOwn</span></span>|<span data-ttu-id="1c780-129">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="1c780-129">16 (010000)</span></span>|<span data-ttu-id="1c780-130">Возможность изменения только созданных ими элементов.</span><span class="sxs-lookup"><span data-stu-id="1c780-130">Can edit only the items they created.</span></span>|
|<span data-ttu-id="1c780-131">едиталл</span><span class="sxs-lookup"><span data-stu-id="1c780-131">EditAll</span></span>|<span data-ttu-id="1c780-132">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="1c780-132">32 (100000)</span></span>|<span data-ttu-id="1c780-133">Можно изменять любые элементы.</span><span class="sxs-lookup"><span data-stu-id="1c780-133">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="1c780-134">В настоящее время API поддерживает доступ к существующим делегированным разрешениям, но не настраивает разрешения делегата.</span><span class="sxs-lookup"><span data-stu-id="1c780-134">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="1c780-135">Объект [делегатепермиссионс](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битовой маски для указания разрешений делегата.</span><span class="sxs-lookup"><span data-stu-id="1c780-135">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="1c780-136">Каждое положение в битовой маске представляет конкретное разрешение и, если ему присвоено значение, `1` у делегата есть соответствующее разрешение.</span><span class="sxs-lookup"><span data-stu-id="1c780-136">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="1c780-137">Например, если второй бит справа `1` , то делегат имеет разрешение на **запись** .</span><span class="sxs-lookup"><span data-stu-id="1c780-137">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="1c780-138">Вы можете увидеть пример того, как проверить наличие определенного разрешения в разделе [выполнение операции как делегата](#perform-an-operation-as-delegate) далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="1c780-138">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="1c780-139">Синхронизация между клиентами почтовых ящиков</span><span class="sxs-lookup"><span data-stu-id="1c780-139">Sync across mailbox clients</span></span>

<span data-ttu-id="1c780-140">Обновление делегата почтового ящика владельца обычно синхронизируется в почтовых ящиках немедленно.</span><span class="sxs-lookup"><span data-stu-id="1c780-140">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="1c780-141">Тем не менее, если для задания расширенного свойства элемента использовались операции REST или Exchange Web Services (EWS), такие изменения могут занять несколько часов. Мы рекомендуем вместо этого использовать объект [CustomProperties](/javascript/api/outlook/office.customproperties) и связанные с ним API, чтобы избежать такой задержки.</span><span class="sxs-lookup"><span data-stu-id="1c780-141">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="1c780-142">Чтобы узнать больше, ознакомьтесь с [разделом Настраиваемые свойства](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи "получение и Настройка метаданных в надстройке Outlook".</span><span class="sxs-lookup"><span data-stu-id="1c780-142">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1c780-143">В сценарии делегата EWS невозможно использовать с маркерами, которые в настоящее время предоставляются office.js API.</span><span class="sxs-lookup"><span data-stu-id="1c780-143">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="1c780-144">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="1c780-144">Configure the manifest</span></span>

<span data-ttu-id="1c780-145">Чтобы включить сценарии делегирования доступа в надстройке, необходимо задать элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` в манифесте под родительским элементом `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="1c780-145">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="1c780-146">В настоящее время другие конструктивные параметры не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="1c780-146">At present, other form factors are not supported.</span></span>

<span data-ttu-id="1c780-147">Чтобы обеспечить поддержку вызовов REST от делегата, задайте для узла [Permissions](../reference/manifest/permissions.md) в манифесте значение `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="1c780-147">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="1c780-148">В приведенном ниже примере показано, как `SupportsSharedFolders` задать элемент `true` в разделе манифеста.</span><span class="sxs-lookup"><span data-stu-id="1c780-148">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="1c780-149">Выполнение операции в качестве делегата</span><span class="sxs-lookup"><span data-stu-id="1c780-149">Perform an operation as delegate</span></span>

<span data-ttu-id="1c780-150">Вы можете получить общие свойства элемента в режиме создания или чтения, вызвав метод [Item. жетшаредпропертиесасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="1c780-150">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="1c780-151">Возвращает объект [шаредпропертиес](/javascript/api/outlook/office.sharedproperties) , который в настоящее время предоставляет разрешения делегата, адрес электронной почты владельца, базовый URL-адрес REST API и целевой почтовый ящик.</span><span class="sxs-lookup"><span data-stu-id="1c780-151">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1c780-152">В сценарии делегирования надстройка может использовать REST, но не EWS, а разрешение надстройки должно быть настроено на разрешение `ReadWriteMailbox` REST доступа к почтовому ящику владельца.</span><span class="sxs-lookup"><span data-stu-id="1c780-152">In a delegate scenario, your add-in can use REST but not EWS, and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox.</span></span>

<span data-ttu-id="1c780-153">В приведенном ниже примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата разрешение на **запись** , и СОВЕРШИТЬ вызов REST.</span><span class="sxs-lookup"><span data-stu-id="1c780-153">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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
> <span data-ttu-id="1c780-154">Как представитель вы можете использовать REST для [получения содержимого сообщения Outlook, присоединенного к элементу Outlook или записи группы](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span><span class="sxs-lookup"><span data-stu-id="1c780-154">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="see-also"></a><span data-ttu-id="1c780-155">См. также</span><span class="sxs-lookup"><span data-stu-id="1c780-155">See also</span></span>

- [<span data-ttu-id="1c780-156">Предоставление другим пользователям возможности управлять почтой и календарем</span><span class="sxs-lookup"><span data-stu-id="1c780-156">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="1c780-157">Общий доступ к календарю в Office 365</span><span class="sxs-lookup"><span data-stu-id="1c780-157">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="1c780-158">Порядок элементов манифеста</span><span class="sxs-lookup"><span data-stu-id="1c780-158">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="1c780-159">[Mask (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="1c780-159">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="1c780-160">Битовые операторы JavaScript</span><span class="sxs-lookup"><span data-stu-id="1c780-160">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)