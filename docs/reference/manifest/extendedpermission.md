---
title: Элемент Екстендедпермиссион в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 138acafb359e2b6e386b34fde7201b1b2c4b3177
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430928"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="fc5f3-103">`ExtendedPermission` элементами</span><span class="sxs-lookup"><span data-stu-id="fc5f3-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="fc5f3-104">Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.</span><span class="sxs-lookup"><span data-stu-id="fc5f3-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="fc5f3-105">`ExtendedPermission`Элемент является дочерним элементом объекта [екстендедпермиссионс](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="fc5f3-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fc5f3-106">Этот элемент доступен только в [предварительной версии требования к надстройке Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="fc5f3-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="fc5f3-107">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="fc5f3-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="fc5f3-108">Доступные расширенные разрешения</span><span class="sxs-lookup"><span data-stu-id="fc5f3-108">Available extended permissions</span></span>

<span data-ttu-id="fc5f3-109">Ниже приведены доступные значения.</span><span class="sxs-lookup"><span data-stu-id="fc5f3-109">The following are the available values.</span></span>

|<span data-ttu-id="fc5f3-110">Доступное значение</span><span class="sxs-lookup"><span data-stu-id="fc5f3-110">Available value</span></span>|<span data-ttu-id="fc5f3-111">Описание</span><span class="sxs-lookup"><span data-stu-id="fc5f3-111">Description</span></span>|<span data-ttu-id="fc5f3-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="fc5f3-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="fc5f3-113">Объявляет, что надстройка использует API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="fc5f3-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="fc5f3-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="fc5f3-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="fc5f3-115">`ExtendedPermission` Примеры</span><span class="sxs-lookup"><span data-stu-id="fc5f3-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="fc5f3-116">Ниже приведен пример `ExtendedPermission` элемента.</span><span class="sxs-lookup"><span data-stu-id="fc5f3-116">The following is an example of the `ExtendedPermission` element.</span></span>

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
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a><span data-ttu-id="fc5f3-117">Содержится в</span><span class="sxs-lookup"><span data-stu-id="fc5f3-117">Contained in</span></span>

[<span data-ttu-id="fc5f3-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="fc5f3-118">ExtendedPermissions</span></span>](extendedpermissions.md)
