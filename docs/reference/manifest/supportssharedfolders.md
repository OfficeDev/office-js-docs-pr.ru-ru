---
title: Элемент SupportsSharedFolders в файле манифеста
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 4ce78d9ece901d8cd6f8639ce7a286f70893a2b4
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120609"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="94a7b-102">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="94a7b-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="94a7b-103">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="94a7b-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="94a7b-104">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="94a7b-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="94a7b-105">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="94a7b-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="94a7b-106">Элемент **SupportsSharedFolders** поддерживается только в Outlook в Интернете и в Windows.</span><span class="sxs-lookup"><span data-stu-id="94a7b-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="94a7b-107">Поддержка этого элемента была введена в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="94a7b-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="94a7b-108">См [клиенты и платформы](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="94a7b-108">See [clients and platforms](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="94a7b-109">Ниже приведен пример элемента  **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="94a7b-109">The following is an example of the  **SupportsSharedFolders** element.</span></span>

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
