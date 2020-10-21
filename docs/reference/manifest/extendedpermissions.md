---
title: Элемент Екстендедпермиссионс в файле манифеста
description: Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626444"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="e4a5a-103">Элемент Екстендедпермиссионс</span><span class="sxs-lookup"><span data-stu-id="e4a5a-103">ExtendedPermissions element</span></span>

<span data-ttu-id="e4a5a-104">Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.</span><span class="sxs-lookup"><span data-stu-id="e4a5a-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="e4a5a-105">`ExtendedPermissions`Элемент является дочерним элементом объекта [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="e4a5a-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e4a5a-106">Поддержка этого элемента была введена в наборе требований 1,9.</span><span class="sxs-lookup"><span data-stu-id="e4a5a-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="e4a5a-107">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="e4a5a-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e4a5a-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e4a5a-108">Child elements</span></span>

|  <span data-ttu-id="e4a5a-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="e4a5a-109">Element</span></span> |  <span data-ttu-id="e4a5a-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e4a5a-110">Required</span></span>  |  <span data-ttu-id="e4a5a-111">Описание</span><span class="sxs-lookup"><span data-stu-id="e4a5a-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="e4a5a-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="e4a5a-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="e4a5a-113">Нет</span><span class="sxs-lookup"><span data-stu-id="e4a5a-113">No</span></span>   | <span data-ttu-id="e4a5a-114">Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.</span><span class="sxs-lookup"><span data-stu-id="e4a5a-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="e4a5a-115">`ExtendedPermissions` Примеры</span><span class="sxs-lookup"><span data-stu-id="e4a5a-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="e4a5a-116">Ниже приведен пример `ExtendedPermissions` элемента.</span><span class="sxs-lookup"><span data-stu-id="e4a5a-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="e4a5a-117">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e4a5a-117">Contained in</span></span>

[<span data-ttu-id="e4a5a-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="e4a5a-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="e4a5a-119">Может содержать</span><span class="sxs-lookup"><span data-stu-id="e4a5a-119">Can contain</span></span>

[<span data-ttu-id="e4a5a-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="e4a5a-120">ExtendedPermission</span></span>](extendedpermission.md)
