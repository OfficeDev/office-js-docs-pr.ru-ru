---
title: Элемент ExtendedPermissions в файле манифеста
description: Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 46ca6e3e2fb992755d9067b4251200073f07ade1
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042128"
---
# <a name="extendedpermissions-element"></a>Элемент ExtendedPermissions

Определяет коллекцию расширенных разрешений, необходимых надстройке для доступа к связанным API или функциям. Элемент `ExtendedPermissions` является детским элементом [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.9. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides:**

- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

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
