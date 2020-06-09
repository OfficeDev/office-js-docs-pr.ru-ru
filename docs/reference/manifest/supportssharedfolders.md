---
title: Элемент SupportsSharedFolders в файле манифеста
description: Элемент SupportsSharedFolders определяет, доступна ли надстройка Outlook в сценариях делегирования.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 3835f7060cc52a72ff0a5ed4dbdb9f1e09258669
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608714"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="8e2ae-103">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="8e2ae-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="8e2ae-104">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="8e2ae-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="8e2ae-105">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="8e2ae-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="8e2ae-106">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="8e2ae-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8e2ae-107">Элемент **SupportsSharedFolders** поддерживается только в Outlook в Интернете и в Windows.</span><span class="sxs-lookup"><span data-stu-id="8e2ae-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="8e2ae-108">Поддержка этого элемента была введена в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="8e2ae-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="8e2ae-109">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="8e2ae-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="8e2ae-110">Ниже приведен пример элемента **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="8e2ae-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
