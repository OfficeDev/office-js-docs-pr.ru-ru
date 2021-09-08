---
title: Элемент ExtendedPermission в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройки для доступа к связанному API или функции.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 37859350cfaffdc14ab91d5026d67aa0a736ac56
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938857"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` элемент

Определяет расширенное разрешение, необходимое надстройки для доступа к связанному API или функции. Элемент `ExtendedPermission` является детским элементом [ExtendedPermissions.](extendedpermissions.md)

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.9. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="available-extended-permissions"></a>Доступные расширенные разрешения

Ниже приводится доступное значение.

|Доступное значение|Описание|Hosts|
|---|---|---|
|`AppendOnSend`|Объявляет, что надстройка использует [Office. API Body.appendOnSendAsync.](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendOnSendAsync_data__options__callback_)|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` пример

Ниже приводится пример `ExtendedPermission` элемента.

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
