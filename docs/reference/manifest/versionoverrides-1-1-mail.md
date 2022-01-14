---
title: VersionOverrides элемент 1.1 в файле манифеста для надстройки почты
description: Справочная документация элемента VersionOverrides 1.1 (почта) для Office файлов манифеста надстройок (XML).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d3187b1c6c60db47e23709f21f264d3c3b0538e4
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042199"
---
# <a name="versionoverrides-11-element-in-the-manifest-file-for-a-mail-add-in"></a>VersionOverrides элемент 1.1 в файле манифеста для надстройки почты

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с обзором элемента [VersionOverrides,](versionoverrides.md)который содержит важную информацию о атрибутах и вариантах элемента.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides:**

- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)
- Некоторые детские элементы могут быть связаны с дополнительными наборами требований.

## <a name="child-elements"></a>Дочерние элементы

В следующей таблице используется только версия 1.1 элементов **VersionOverrides** и только почтовые надстройки.

> [!NOTE]
> Только в iOS `<WebApplicationInfo>` поддерживается. Все остальные детские элементы **VersionOverrides** игнорируются.

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Описание](#description)    |  Нет   |  Описывает надстройку. |
|  [Requirements](requirements.md)  |  Нет   |  Указывает минимальные наборы требований, которые необходимо поддерживать для того, чтобы разметка в родительском документе `VersionOverrides` вступила в силу. Это всегда должно быть *более строгим,* чем элемент `Requirements` в базовой части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Указывает коллекцию Office приложений. Элемент Child Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](equivalentaddins.md)    |  Нет  | Указывает родной (COM/XLL) надстройки, эквивалентные веб-надстройки. Веб-надстройка не активируется, если установлена эквивалентная родной надстройка.|
|  **VersionOverrides**    |  Нет  | В настоящее время не может быть в VersionOverrides 1.1 для почтовых надстройок. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки с защищенными эмитентами маркеров, такими как Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Нет  |  Указывает коллекцию расширенных разрешений. |

### <a name="description"></a>Описание

Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Атрибут элемента Description может быть не более 32 символов и за набором значения атрибута `resid`  `id` `String` элемента, содержаного текст.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [AddinCommands 1.1,](../requirement-sets/add-in-commands-requirement-sets.md) когда родительский `<VersionOverrides>` тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский `<VersionOverrides>` тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский `<VersionOverrides>` тип Почта 1.1.

## <a name="example"></a>Пример

Ниже приведен простой пример. Более полные примеры см. в манифестах для примеров надстройок в Office примерах [кода надстройки.](https://github.com/OfficeDev/PnP-OfficeAddins)

Ниже приводится пример типичного элемента, включая некоторые детские элементы, которые не требуются, но `<VersionOverrides>` обычно используются.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Реализация нескольких версий

В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.

Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии. Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.

Чтобы реализовать схему VersionOverrides v1.0 и v1.1, манифест будет выглядеть аналогично следующему примеру.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
