---
title: Элемент Екстендедпермиссионс в файле манифеста
description: Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 86d898052af6ba0e6f6bc8b341fff9f0f8408967
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718225"
---
# <a name="extendedpermissions-element"></a>Элемент Екстендедпермиссионс

Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям. `ExtendedPermissions` Элемент является дочерним элементом объекта [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Этот элемент доступен только в [предварительной версии требования к надстройке Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online. Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Нет   | Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions`Примеры

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
