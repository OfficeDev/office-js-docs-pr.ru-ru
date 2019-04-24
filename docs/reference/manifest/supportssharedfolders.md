---
title: Элемент SupportsSharedFolders в файле манифеста
description: ''
ms.date: 04/02/2019
localization_priority: Normal
ms.openlocfilehash: 976f8ba00f6ac9ac32def56933af1077527b7e9c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452041"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="3b8e6-102">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="3b8e6-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="3b8e6-103">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="3b8e6-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="3b8e6-104">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="3b8e6-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="3b8e6-105">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="3b8e6-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3b8e6-106">Доступ представителей для надстроек Outlook в настоящее время находится [в предварительной версии](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) и поддерживается только в клиентах, работающих в Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="3b8e6-106">Delegate access for Outlook add-ins is currently [in preview](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="3b8e6-107">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="3b8e6-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="3b8e6-108">Ниже приведен пример элемента  **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="3b8e6-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

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
