---
title: Элемент SupportsSharedFolders в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 776d44ec66c4e27a72e5487051bed1edf4b3dcaf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432685"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="fb30d-102">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="fb30d-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="fb30d-103">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="fb30d-103">It defines whether the add-in is available in delegate scenarios.</span></span> <span data-ttu-id="fb30d-104">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="fb30d-104">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="fb30d-105">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="fb30d-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fb30d-106">Этот элемент доступен только в [предварительной версии набора обязательных элементов надстроек Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="fb30d-106">This element is only available in the [Outlook add-ins Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="fb30d-107">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="fb30d-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="fb30d-108">Ниже приведен пример элемента  **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="fb30d-108">The following is an example of how the **Rows** element should look.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
