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
# <a name="supportssharedfolders-element"></a>Элемент SupportsSharedFolders

Определяет, доступна ли надстройка Outlook в сценариях делегирования. Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md). По умолчанию для него установлено значение *false*.

> [!IMPORTANT]
> Элемент **SupportsSharedFolders** поддерживается только в Outlook в Интернете и в Windows.
>
> Поддержка этого элемента была введена в наборе требований 1,8. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

Ниже приведен пример элемента **SupportsSharedFolders** .

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
