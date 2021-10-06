---
title: Элемент SupportsSharedFolders в файле манифеста
description: Элемент SupportsSharedFolders определяет, доступна ли надстройка Outlook в общих папках и сценариях общих почтовых ящиков.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e13393f10b12e0a3c5ca1b004b202eb2970d264
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138725"
---
# <a name="supportssharedfolders-element"></a>Элемент SupportsSharedFolders

Определяет, доступна ли надстройка Outlook в общих почтовых ящиках (в настоящее время в предварительном просмотре) и общих папках (т. е. в сценариях делегирования доступа). Элемент **SupportsSharedFolders** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md). По умолчанию для него установлено значение *false*.

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.8. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides:**

- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [Mailbox 1.8](../../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)

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
