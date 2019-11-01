---
title: Элемент SupportsSharedFolders в файле манифеста
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 42fa1cf74634b183994e633d728d3be66e1e83f0
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902244"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="e32e6-102">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e32e6-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="e32e6-103">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="e32e6-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="e32e6-104">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e32e6-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="e32e6-105">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="e32e6-105">It is set to *false* by default.</span></span>

<span data-ttu-id="e32e6-106">Ниже приведен пример элемента  **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="e32e6-106">The following is an example of the  **SupportsSharedFolders** element.</span></span>

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
  </VersionOverrides>
</VersionOverrides>
...
```
