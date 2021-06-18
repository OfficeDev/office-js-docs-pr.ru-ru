---
title: Элемент SupportsSharedFolders в файле манифеста
description: Элемент SupportsSharedFolders определяет, доступна ли надстройка Outlook в общих папках и сценариях общих почтовых ящиков.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 43f2c60664a6822b714023246cfa044e179e9a55
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007785"
---
# <a name="supportssharedfolders-element"></a>Элемент SupportsSharedFolders

Определяет, доступна ли надстройка Outlook в общих почтовых ящиках (в настоящее время в предварительном просмотре) и общих папках (т. е. в сценариях делегирования доступа). Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md). По умолчанию для него установлено значение *false*.

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.8. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

Ниже приводится пример элемента **SupportsSharedFolders.**

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
