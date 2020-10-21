---
title: Элемент Екстендедпермиссион в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626402"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` элементами

Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции. `ExtendedPermission`Элемент является дочерним элементом объекта [екстендедпермиссионс](extendedpermissions.md).

> [!IMPORTANT]
> Поддержка этого элемента была введена в наборе требований 1,9. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="available-extended-permissions"></a>Доступные расширенные разрешения

Ниже приведены доступные значения.

|Доступное значение|Описание|Hosts|
|---|---|---|
|`AppendOnSend`|Объявляет, что надстройка использует API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` Примеры

Ниже приведен пример `ExtendedPermission` элемента.

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

[ExtendedPermissions](extendedpermissions.md)
