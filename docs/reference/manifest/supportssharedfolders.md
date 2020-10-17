---
title: Элемент SupportsSharedFolders в файле манифеста
description: Элемент SupportsSharedFolders определяет, доступна ли надстройка Outlook в сценариях делегирования.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 786a4763450d78cb16c9baafc81701758af54787
ms.sourcegitcommit: 6fa29989dfaec4dfa0f8df3fe5fb038d7afbae30
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/16/2020
ms.locfileid: "48487882"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="35fd4-103">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="35fd4-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="35fd4-104">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="35fd4-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="35fd4-105">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="35fd4-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="35fd4-106">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="35fd4-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="35fd4-107">Поддержка этого элемента была введена в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="35fd4-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="35fd4-108">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="35fd4-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="35fd4-109">Ниже приведен пример элемента **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="35fd4-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
