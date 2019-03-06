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
# <a name="supportssharedfolders-element"></a>Элемент SupportsSharedFolders

Определяет, доступна ли надстройка Outlook в сценариях делегирования. Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md). По умолчанию для него установлено значение *false*.

> [!IMPORTANT]
> Доступ представителей для надстроек Outlook в настоящее время находится в предварительной версии и поддерживается только в клиентах, работающих в Exchange Online. Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.

Ниже приведен пример элемента  **SupportsSharedFolders**.

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
