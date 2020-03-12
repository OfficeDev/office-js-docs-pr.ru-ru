---
title: Элемент SupportsSharedFolders в файле манифеста
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 81401b79f4c443305e376df7a66a07d916393d17
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596755"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="e5b25-102">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e5b25-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="e5b25-103">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="e5b25-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="e5b25-104">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e5b25-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="e5b25-105">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="e5b25-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e5b25-106">Элемент **SupportsSharedFolders** поддерживается только в Outlook в Интернете и в Windows.</span><span class="sxs-lookup"><span data-stu-id="e5b25-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="e5b25-107">Поддержка этого элемента была введена в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="e5b25-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="e5b25-108">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="e5b25-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="e5b25-109">Ниже приведен пример элемента **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="e5b25-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
