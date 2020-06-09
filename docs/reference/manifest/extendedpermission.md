---
title: Элемент Екстендедпермиссион в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: ca4c2d12cb9a5be159c22712b631d0bde42e48ed
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611542"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission`элементами

Определяет расширенное разрешение, необходимое надстройке для доступа к связанному API или функции. `ExtendedPermission`Элемент является дочерним элементом объекта [екстендедпермиссионс](extendedpermissions.md).

> [!IMPORTANT]
> Этот элемент доступен только в [предварительной версии требования к надстройке Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) для Exchange Online. Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.

## <a name="available-extended-permissions"></a>Доступные расширенные разрешения

Ниже приведены доступные значения.

|Доступное значение|Description|Hosts|
|---|---|---|
|`AppendOnSend`|Объявляет, что надстройка использует API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission`Примеры

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
