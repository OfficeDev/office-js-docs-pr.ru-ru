---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация элемента VersionOverrides для Office файлов манифеста надстройок (XML).
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 787ba8e7d90900cc72d6c5e9370d68ced0faee2f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936530"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Пространство имен VersionOverrides схемы. Допустимые значения различаются в зависимости от `<VersionOverrides>` **значения xsi:type** этого элемента и **значения xsi:type** родительского `<OfficeApp>` элемента. Ниже [см. значения Пространства](#namespace-values) имен.|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространства имен

Ниже перечислены необходимое значение **значения xmlns** в зависимости от **значения xsi:type** родительского `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns должны** быть `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns должны** быть `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** поддерживает версии 1.0 и 1.1 VersionOverrides, поэтому значение **xmlns** зависит от значения `<VersionOverrides>` **xsi:type** этого элемента:
    - Когда **xsi:type** `VersionOverridesV1_0` , **xmlns должны** быть `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Когда **xsi:type** `VersionOverridesV1_1` , **xmlns должны** быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> В настоящее время Outlook 2016 поддерживает схему VersionOverrides v1.1 и `VersionOverridesV1_1` тип.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  Нет   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Атрибут элемента Description может быть не более 32 символов и за набором значения атрибута `resid`  `id` `String` элемента, содержаного текст.|
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Указывает коллекцию Office приложений. Элемент Child Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](equivalentaddins.md)    |  Нет  | Указывает родной (COM/XLL) надстройки, эквивалентные веб-надстройки. Веб-надстройка не активируется, если установлена эквивалентная родной надстройка.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки с защищенными эмитентами маркеров, такими как Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Нет  |  Указывает коллекцию расширенных разрешений. |

### <a name="versionoverrides-example"></a>Пример VersionOverrides

Ниже приводится пример типичного элемента, включая некоторые детские элементы, которые не требуются, но `<VersionOverrides>` обычно используются.

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
