---
title: Элемент Екстендедпермиссионс в файле манифеста
description: Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626444"
---
# <a name="extendedpermissions-element"></a>Элемент Екстендедпермиссионс

Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям. `ExtendedPermissions`Элемент является дочерним элементом объекта [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Поддержка этого элемента была введена в наборе требований 1,9. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Нет   | Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` Примеры

Ниже приведен пример `ExtendedPermissions` элемента.

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
