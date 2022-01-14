---
title: Элемент ExtendedPermission в файле манифеста
description: Определяет расширенное разрешение, необходимое надстройки для доступа к связанному API или функции.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5ed3745da87c2fa04839a8fbd1c677f62ad771dc
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042142"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` элемент

Определяет расширенное разрешение, необходимое надстройки для доступа к связанному API или функции. Элемент `ExtendedPermission` является детским элементом [ExtendedPermissions.](extendedpermissions.md)

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований 1.9. См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides:**

- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

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
