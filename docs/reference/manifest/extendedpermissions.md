---
title: Элемент ExtendedPermissions в файле манифеста
description: Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938849"
---
# <a name="extendedpermissions-element"></a>Элемент ExtendedPermissions

Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям. Элемент `ExtendedPermissions` является детским элементом [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.9. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Нет   | Определяет расширенное разрешение, необходимое для надстройки для доступа к связанному API или функции. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` пример

Ниже приводится пример `ExtendedPermissions` элемента.

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
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>Содержится в

[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>Может содержать

[ExtendedPermission](extendedpermission.md)
