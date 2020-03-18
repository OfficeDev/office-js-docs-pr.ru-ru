---
title: Элемент Екстендедпермиссионс в файле манифеста
description: Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 86d898052af6ba0e6f6bc8b341fff9f0f8408967
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718225"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="6838d-103">Элемент Екстендедпермиссионс</span><span class="sxs-lookup"><span data-stu-id="6838d-103">ExtendedPermissions element</span></span>

<span data-ttu-id="6838d-104">Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.</span><span class="sxs-lookup"><span data-stu-id="6838d-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="6838d-105">`ExtendedPermissions` Элемент является дочерним элементом объекта [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="6838d-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6838d-106">Этот элемент доступен только в [предварительной версии требования к надстройке Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="6838d-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="6838d-107">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="6838d-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6838d-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6838d-108">Child elements</span></span>

|  <span data-ttu-id="6838d-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="6838d-109">Element</span></span> |  <span data-ttu-id="6838d-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6838d-110">Required</span></span>  |  <span data-ttu-id="6838d-111">Описание</span><span class="sxs-lookup"><span data-stu-id="6838d-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="6838d-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="6838d-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="6838d-113">Нет</span><span class="sxs-lookup"><span data-stu-id="6838d-113">No</span></span>   | <span data-ttu-id="6838d-114">Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.</span><span class="sxs-lookup"><span data-stu-id="6838d-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="6838d-115">`ExtendedPermissions`Примеры</span><span class="sxs-lookup"><span data-stu-id="6838d-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="6838d-116">Ниже приведен пример `ExtendedPermissions` элемента.</span><span class="sxs-lookup"><span data-stu-id="6838d-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="6838d-117">Содержится в</span><span class="sxs-lookup"><span data-stu-id="6838d-117">Contained in</span></span>

[<span data-ttu-id="6838d-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="6838d-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="6838d-119">Может содержать</span><span class="sxs-lookup"><span data-stu-id="6838d-119">Can contain</span></span>

[<span data-ttu-id="6838d-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="6838d-120">ExtendedPermission</span></span>](extendedpermission.md)
