---
title: Включение сценариев делегирования доступа в надстройке Outlook
description: В кратко описывается доступ представителя и описывается настройка поддержки надстройки.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 68b9e09afbe2bcd5cfc302d6714b1c22fd945047
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608952"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="8db4f-103">Включение сценариев делегирования доступа в надстройке Outlook</span><span class="sxs-lookup"><span data-stu-id="8db4f-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="8db4f-104">Владелец почтового ящика может использовать функцию делегированного доступа, чтобы [Разрешить другому пользователю управлять своей почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span><span class="sxs-lookup"><span data-stu-id="8db4f-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="8db4f-105">В этой статье указывается, какие разрешения представителей поддерживает API JavaScript для Office, а также описывается включение сценариев делегированного доступа в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="8db4f-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8db4f-106">Доступ к представителю в настоящее время недоступен в Outlook для Mac, Android и iOS.</span><span class="sxs-lookup"><span data-stu-id="8db4f-106">Delegate access is not currently available in Outlook on Mac, Android, and iOS.</span></span> <span data-ttu-id="8db4f-107">Эта функция может быть доступна в будущем.</span><span class="sxs-lookup"><span data-stu-id="8db4f-107">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="8db4f-108">Поддержка этой функции появилась в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="8db4f-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="8db4f-109">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="8db4f-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="8db4f-110">Поддерживаемые разрешения для делегированного доступа</span><span class="sxs-lookup"><span data-stu-id="8db4f-110">Supported permissions for delegate access</span></span>

<span data-ttu-id="8db4f-111">В следующей таблице описаны разрешения представителей, поддерживаемые API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="8db4f-111">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="8db4f-112">Permission</span><span class="sxs-lookup"><span data-stu-id="8db4f-112">Permission</span></span>|<span data-ttu-id="8db4f-113">Значение</span><span class="sxs-lookup"><span data-stu-id="8db4f-113">Value</span></span>|<span data-ttu-id="8db4f-114">Описание</span><span class="sxs-lookup"><span data-stu-id="8db4f-114">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="8db4f-115">Read</span><span class="sxs-lookup"><span data-stu-id="8db4f-115">Read</span></span>|<span data-ttu-id="8db4f-116">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="8db4f-116">1 (000001)</span></span>|<span data-ttu-id="8db4f-117">Возможность чтения элементов.</span><span class="sxs-lookup"><span data-stu-id="8db4f-117">Can read items.</span></span>|
|<span data-ttu-id="8db4f-118">Write</span><span class="sxs-lookup"><span data-stu-id="8db4f-118">Write</span></span>|<span data-ttu-id="8db4f-119">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="8db4f-119">2 (000010)</span></span>|<span data-ttu-id="8db4f-120">Может создавать элементы.</span><span class="sxs-lookup"><span data-stu-id="8db4f-120">Can create items.</span></span>|
|<span data-ttu-id="8db4f-121">делетеовн</span><span class="sxs-lookup"><span data-stu-id="8db4f-121">DeleteOwn</span></span>|<span data-ttu-id="8db4f-122">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="8db4f-122">4 (000100)</span></span>|<span data-ttu-id="8db4f-123">Можно удалять только созданные ими элементы.</span><span class="sxs-lookup"><span data-stu-id="8db4f-123">Can delete only the items they created.</span></span>|
|<span data-ttu-id="8db4f-124">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="8db4f-124">DeleteAll</span></span>|<span data-ttu-id="8db4f-125">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="8db4f-125">8 (001000)</span></span>|<span data-ttu-id="8db4f-126">Может удалять все элементы.</span><span class="sxs-lookup"><span data-stu-id="8db4f-126">Can delete any items.</span></span>|
|<span data-ttu-id="8db4f-127">едитовн</span><span class="sxs-lookup"><span data-stu-id="8db4f-127">EditOwn</span></span>|<span data-ttu-id="8db4f-128">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="8db4f-128">16 (010000)</span></span>|<span data-ttu-id="8db4f-129">Возможность изменения только созданных ими элементов.</span><span class="sxs-lookup"><span data-stu-id="8db4f-129">Can edit only the items they created.</span></span>|
|<span data-ttu-id="8db4f-130">едиталл</span><span class="sxs-lookup"><span data-stu-id="8db4f-130">EditAll</span></span>|<span data-ttu-id="8db4f-131">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="8db4f-131">32 (100000)</span></span>|<span data-ttu-id="8db4f-132">Можно изменять любые элементы.</span><span class="sxs-lookup"><span data-stu-id="8db4f-132">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="8db4f-133">В настоящее время API поддерживает доступ к существующим делегированным разрешениям, но не настраивает разрешения делегата.</span><span class="sxs-lookup"><span data-stu-id="8db4f-133">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="8db4f-134">Объект [делегатепермиссионс](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битовой маски для указания разрешений делегата.</span><span class="sxs-lookup"><span data-stu-id="8db4f-134">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="8db4f-135">Каждое положение в битовой маске представляет конкретное разрешение и, если ему присвоено значение, `1` у делегата есть соответствующее разрешение.</span><span class="sxs-lookup"><span data-stu-id="8db4f-135">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="8db4f-136">Например, если второй бит справа `1` , то делегат имеет разрешение на **запись** .</span><span class="sxs-lookup"><span data-stu-id="8db4f-136">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="8db4f-137">Вы можете увидеть пример того, как проверить наличие определенного разрешения в разделе [выполнение операции как делегата](#perform-an-operation-as-delegate) далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="8db4f-137">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="8db4f-138">Синхронизация между клиентами почтовых ящиков</span><span class="sxs-lookup"><span data-stu-id="8db4f-138">Sync across mailbox clients</span></span>

<span data-ttu-id="8db4f-139">Обновление делегата почтового ящика владельца обычно синхронизируется в почтовых ящиках немедленно.</span><span class="sxs-lookup"><span data-stu-id="8db4f-139">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="8db4f-140">Тем не менее, если надстройка использует операции REST или EWS для задания расширенного свойства элемента, такие изменения могут занять несколько часов. Мы рекомендуем вместо этого использовать объект [CustomProperties](/javascript/api/outlook/office.customproperties) и связанные с ним API, чтобы избежать такой задержки.</span><span class="sxs-lookup"><span data-stu-id="8db4f-140">However, if the add-in uses REST or EWS operations to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="8db4f-141">Чтобы узнать больше, ознакомьтесь с [разделом Настраиваемые свойства](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи "получение и Настройка метаданных в надстройке Outlook".</span><span class="sxs-lookup"><span data-stu-id="8db4f-141">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="8db4f-142">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="8db4f-142">Configure the manifest</span></span>

<span data-ttu-id="8db4f-143">Чтобы включить сценарии делегирования доступа в надстройке, необходимо задать элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` в манифесте под родительским элементом `DesktopFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="8db4f-143">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="8db4f-144">В настоящее время другие конструктивные параметры не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="8db4f-144">At present, other form factors are not supported.</span></span>

<span data-ttu-id="8db4f-145">В приведенном ниже примере показано, как `SupportsSharedFolders` задать элемент `true` в разделе манифеста.</span><span class="sxs-lookup"><span data-stu-id="8db4f-145">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

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

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="8db4f-146">Выполнение операции в качестве делегата</span><span class="sxs-lookup"><span data-stu-id="8db4f-146">Perform an operation as delegate</span></span>

<span data-ttu-id="8db4f-147">Вы можете получить общие свойства элемента в режиме создания или чтения, вызвав метод [Item. жетшаредпропертиесасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="8db4f-147">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="8db4f-148">Возвращает объект [шаредпропертиес](/javascript/api/outlook/office.sharedproperties) , который в настоящее время предоставляет разрешения делегата, адрес электронной почты владельца, базовый URL-адрес REST API и целевой почтовый ящик.</span><span class="sxs-lookup"><span data-stu-id="8db4f-148">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="8db4f-149">В приведенном ниже примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата разрешение на **запись** , и СОВЕРШИТЬ вызов REST.</span><span class="sxs-lookup"><span data-stu-id="8db4f-149">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="8db4f-150">См. также</span><span class="sxs-lookup"><span data-stu-id="8db4f-150">See also</span></span>

- [<span data-ttu-id="8db4f-151">Предоставление другим пользователям возможности управлять почтой и календарем</span><span class="sxs-lookup"><span data-stu-id="8db4f-151">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="8db4f-152">Общий доступ к календарю в Office 365</span><span class="sxs-lookup"><span data-stu-id="8db4f-152">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="8db4f-153">Порядок элементов манифеста</span><span class="sxs-lookup"><span data-stu-id="8db4f-153">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="8db4f-154">[Mask (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="8db4f-154">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="8db4f-155">Битовые операторы JavaScript</span><span class="sxs-lookup"><span data-stu-id="8db4f-155">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)