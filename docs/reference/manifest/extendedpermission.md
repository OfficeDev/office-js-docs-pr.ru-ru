---
title: Элемент Екстендедпермиссион в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626402"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="f84d2-103">`ExtendedPermission` элементами</span><span class="sxs-lookup"><span data-stu-id="f84d2-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="f84d2-104">Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.</span><span class="sxs-lookup"><span data-stu-id="f84d2-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="f84d2-105">`ExtendedPermission`Элемент является дочерним элементом объекта [екстендедпермиссионс](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="f84d2-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f84d2-106">Поддержка этого элемента была введена в наборе требований 1,9.</span><span class="sxs-lookup"><span data-stu-id="f84d2-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="f84d2-107">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="f84d2-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="f84d2-108">Доступные расширенные разрешения</span><span class="sxs-lookup"><span data-stu-id="f84d2-108">Available extended permissions</span></span>

<span data-ttu-id="f84d2-109">Ниже приведены доступные значения.</span><span class="sxs-lookup"><span data-stu-id="f84d2-109">The following are the available values.</span></span>

|<span data-ttu-id="f84d2-110">Доступное значение</span><span class="sxs-lookup"><span data-stu-id="f84d2-110">Available value</span></span>|<span data-ttu-id="f84d2-111">Описание</span><span class="sxs-lookup"><span data-stu-id="f84d2-111">Description</span></span>|<span data-ttu-id="f84d2-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="f84d2-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="f84d2-113">Объявляет, что надстройка использует API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="f84d2-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="f84d2-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="f84d2-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="f84d2-115">`ExtendedPermission` Примеры</span><span class="sxs-lookup"><span data-stu-id="f84d2-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="f84d2-116">Ниже приведен пример `ExtendedPermission` элемента.</span><span class="sxs-lookup"><span data-stu-id="f84d2-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f84d2-117">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f84d2-117">Contained in</span></span>

[<span data-ttu-id="f84d2-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="f84d2-118">ExtendedPermissions</span></span>](extendedpermissions.md)
