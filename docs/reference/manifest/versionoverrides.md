---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация по элементу VersionOverrides для файлов манифеста надстроек Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 979a75c3ea8b4d600a2c43fc4edfcb0d4e96930e
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431544"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Пространство имен схемы VersionOverrides. Допустимые значения зависят от `<VersionOverrides>` значения **xsi: Type** этого элемента и значения **xsi: Type** родительского `<OfficeApp>` элемента. Ниже приведены [значения пространств имен](#namespace-values) .|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространств имен

Ниже приведен список требуемого значения **xmlns** в зависимости от значения **xsi: Type** родительского `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1,0 VersionOverrides, а **xmlns** — значение `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** поддерживает только версию 1,0 VersionOverrides, а **xmlns** — значение `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** поддерживает версии 1,0 и 1,1 для VersionOverrides, поэтому значение **xmlns** зависит от `<VersionOverrides>` значения **xsi: Type** этого элемента:
    - Если **xsi: Type** `VersionOverridesV1_0` , то **xmlns** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Если **xsi: Type** `VersionOverridesV1_1` , то **xmlns** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> В настоящее время только Outlook 2016 или более поздней версии поддерживает схему VersionOverrides 1.1 и `VersionOverridesV1_1` тип.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  Нет   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.|
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Задает коллекцию приложений Office. Дочерний элемент hosts переопределяет элемент hosts в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](equivalentaddins.md)    |  Нет  | Задает встроенные надстройки (COM/XLL), эквивалентные веб-надстройке. Веб-надстройка не активируется, если установлена эквивалентная собственная встроенная надстройка.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Задает сведения о регистрации надстройки с помощью надежных поставщиков маркеров, таких как Azure Active Directory 2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Нет  |  Задает коллекцию расширенных разрешений.<br><br>**Важно!** поскольку API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) в настоящее время находится в режиме предварительной версии, надстройки, использующие этот `ExtendedPermissions` элемент, не могут быть опубликованы в AppSource или развернуты с помощью централизованного развертывания. |

### <a name="versionoverrides-example"></a>Пример VersionOverrides

Ниже приведен пример типичного `<VersionOverrides>` элемента, в том числе некоторые необязательные дочерние элементы, которые обычно используются.

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
