---
title: Элемент SupportsSharedFolders в файле манифеста
description: Элемент SupportsSharedFolders определяет, доступна ли надстройка Outlook в сценариях делегирования.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 66a426b0c31bda61feb23cb83d63722898dfb503
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717889"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="8a38a-103">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="8a38a-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="8a38a-104">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="8a38a-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="8a38a-105">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="8a38a-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="8a38a-106">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="8a38a-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8a38a-107">Элемент **SupportsSharedFolders** поддерживается только в Outlook в Интернете и в Windows.</span><span class="sxs-lookup"><span data-stu-id="8a38a-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="8a38a-108">Поддержка этого элемента была введена в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="8a38a-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="8a38a-109">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="8a38a-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="8a38a-110">Ниже приведен пример элемента **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="8a38a-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
