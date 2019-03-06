---
title: Элемент SupportsSharedFolders в файле манифеста
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: bfbce42c7d1aa5eefab40b528c5b622aa7d2d54f
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413617"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="505f3-102">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="505f3-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="505f3-103">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="505f3-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="505f3-104">Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="505f3-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="505f3-105">По умолчанию для него установлено значение *false*.</span><span class="sxs-lookup"><span data-stu-id="505f3-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="505f3-106">Доступ представителей для надстроек Outlook в настоящее время находится в предварительной версии и поддерживается только в клиентах, работающих в Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="505f3-106">Delegate access for Outlook add-ins is currently in preview and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="505f3-107">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="505f3-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="505f3-108">Ниже приведен пример элемента  **SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="505f3-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
