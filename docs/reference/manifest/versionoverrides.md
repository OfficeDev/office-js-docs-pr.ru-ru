---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация элемента VersionOverrides для Office дополнительных дополнительных виленок (XML).
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 0a70ded82b4603b1ac70698947a4710a4a44b5b6
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555152"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  ВерсияОвергорайды схема пространства имен. Разрешенные значения варьируются в `<VersionOverrides>` зависимости от **значения xsi:типа** этого элемента **и значения xsi:типа** родительского `<OfficeApp>` элемента. Ниже [приведены значения пространства имен.](#namespace-values)|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространства имен

Ниже приводится перечне требуемое значение **значения xmlns** в зависимости **от значения xsi:type** родительского `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны `http://schemas.microsoft.com/office/taskpaneappversionoverrides` быть.
- **ContentApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны `http://schemas.microsoft.com/office/contentappversionoverrides` быть.
- **MailApp** поддерживает версии 1.0 и 1.1 VersionOverrides, поэтому **значение xmlns варьируется** в `<VersionOverrides>` зависимости от **значения xsi:type** этого элемента:
    - Когда **xsi:type** `VersionOverridesV1_0` есть, **xmlns** должен `http://schemas.microsoft.com/office/mailappversionoverrides` быть.
    - Когда **xsi:type** `VersionOverridesV1_1` есть, **xmlns** должен `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` быть.

> [!NOTE]
> В настоящее Outlook 2016 или позже поддерживает схему VersionOverrides v1.1 и `VersionOverridesV1_1` тип.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  Нет   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Атрибут `resid` элемента **Описание может** быть не более 32 символов и устанавливается на `id` значение атрибута `String` элемента, который содержит текст.|
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Определяет набор Office приложений. Элемент «Хосты ребенка» перекрывает элемент «Хозяева» в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](equivalentaddins.md)    |  Нет  | Определяет родные (COM/XLL) дополнения, эквивалентные веб-надстройки. Веб-надстройок не активируется, если установлена эквивалентная пристройная система.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Уточняется подробная информация о регистрации надстройки с защищенными эмитентами токенов, такими как Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Нет  |  Определяет набор расширенных разрешений. |

### <a name="versionoverrides-example"></a>Пример VersionOverrides

Ниже приводится пример типичного `<VersionOverrides>` элемента, включая некоторые элементы ребенка, которые не требуются, но обычно используются.

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
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Реализация нескольких версий

В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.

Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии. Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.

Чтобы реализовать схему VersionOverrides версий 1.0 и 1.1, манифест должен выглядеть следующим образом:

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
