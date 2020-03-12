---
title: Элемент Екстендедпермиссион в файле манифеста
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 6c41684fc922f5845559250311edd8182788cfc5
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605811"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="a8daf-102">`ExtendedPermission`элементами</span><span class="sxs-lookup"><span data-stu-id="a8daf-102">`ExtendedPermission` element</span></span>

<span data-ttu-id="a8daf-103">Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.</span><span class="sxs-lookup"><span data-stu-id="a8daf-103">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="a8daf-104">`ExtendedPermission` Элемент является дочерним элементом объекта [екстендедпермиссионс](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="a8daf-104">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a8daf-105">Этот элемент доступен только в [предварительной версии требования к надстройке Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="a8daf-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="a8daf-106">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="a8daf-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="a8daf-107">Доступные расширенные разрешения</span><span class="sxs-lookup"><span data-stu-id="a8daf-107">Available extended permissions</span></span>

<span data-ttu-id="a8daf-108">Ниже приведены доступные значения.</span><span class="sxs-lookup"><span data-stu-id="a8daf-108">The following are the available values.</span></span>

|<span data-ttu-id="a8daf-109">Доступное значение</span><span class="sxs-lookup"><span data-stu-id="a8daf-109">Available value</span></span>|<span data-ttu-id="a8daf-110">Описание</span><span class="sxs-lookup"><span data-stu-id="a8daf-110">Description</span></span>|<span data-ttu-id="a8daf-111">Hosts</span><span class="sxs-lookup"><span data-stu-id="a8daf-111">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="a8daf-112">Объявляет, что надстройка использует API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="a8daf-112">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="a8daf-113">Outlook</span><span class="sxs-lookup"><span data-stu-id="a8daf-113">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="a8daf-114">`ExtendedPermission`Примеры</span><span class="sxs-lookup"><span data-stu-id="a8daf-114">`ExtendedPermission` example</span></span>

<span data-ttu-id="a8daf-115">Ниже приведен пример `ExtendedPermission` элемента.</span><span class="sxs-lookup"><span data-stu-id="a8daf-115">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="a8daf-116">Содержится в</span><span class="sxs-lookup"><span data-stu-id="a8daf-116">Contained in</span></span>

[<span data-ttu-id="a8daf-117">екстендедпермиссионс</span><span class="sxs-lookup"><span data-stu-id="a8daf-117">ExtendedPermissions</span></span>](extendedpermissions.md)
