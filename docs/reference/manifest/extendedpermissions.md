---
title: Элемент Екстендедпермиссионс в файле манифеста
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 966378b8bbed66960d7a99c4a82df75ace1c9161
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605814"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="c36cc-102">Элемент Екстендедпермиссионс</span><span class="sxs-lookup"><span data-stu-id="c36cc-102">ExtendedPermissions element</span></span>

<span data-ttu-id="c36cc-103">Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.</span><span class="sxs-lookup"><span data-stu-id="c36cc-103">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="c36cc-104">`ExtendedPermissions` Элемент является дочерним элементом объекта [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="c36cc-104">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c36cc-105">Этот элемент доступен только в [предварительной версии требования к надстройке Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="c36cc-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="c36cc-106">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="c36cc-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c36cc-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c36cc-107">Child elements</span></span>

|  <span data-ttu-id="c36cc-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="c36cc-108">Element</span></span> |  <span data-ttu-id="c36cc-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c36cc-109">Required</span></span>  |  <span data-ttu-id="c36cc-110">Описание</span><span class="sxs-lookup"><span data-stu-id="c36cc-110">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="c36cc-111">екстендедпермиссион</span><span class="sxs-lookup"><span data-stu-id="c36cc-111">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="c36cc-112">Нет</span><span class="sxs-lookup"><span data-stu-id="c36cc-112">No</span></span>   | <span data-ttu-id="c36cc-113">Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.</span><span class="sxs-lookup"><span data-stu-id="c36cc-113">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="c36cc-114">`ExtendedPermissions`Примеры</span><span class="sxs-lookup"><span data-stu-id="c36cc-114">`ExtendedPermissions` example</span></span>

<span data-ttu-id="c36cc-115">Ниже приведен пример `ExtendedPermissions` элемента.</span><span class="sxs-lookup"><span data-stu-id="c36cc-115">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="c36cc-116">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c36cc-116">Contained in</span></span>

[<span data-ttu-id="c36cc-117">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c36cc-117">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="c36cc-118">Может содержать</span><span class="sxs-lookup"><span data-stu-id="c36cc-118">Can contain</span></span>

[<span data-ttu-id="c36cc-119">екстендедпермиссион</span><span class="sxs-lookup"><span data-stu-id="c36cc-119">ExtendedPermission</span></span>](extendedpermission.md)
