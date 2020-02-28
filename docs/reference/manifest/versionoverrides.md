---
title: Элемент VersionOverrides в файле манифеста
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 0afa3183e34a736a878217c079b7b8d0259be5b1
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324801"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Пространство имен схемы VersionOverrides. Допустимые значения зависят от значения `<VersionOverrides>` **xsi: Type** этого элемента и значения **xsi: Type** родительского `<OfficeApp>` элемента. Ниже приведены [значения пространств имен](#namespace-values) .|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространств имен

Ниже приведен список требуемого значения **xmlns** в зависимости от значения **xsi: Type** родительского `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1,0 VersionOverrides, а **xmlns** — значение `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.
- **ContentApp** поддерживает только версию 1,0 VersionOverrides, а **xmlns** — значение `http://schemas.microsoft.com/office/contentappversionoverrides`.
- **MailApp** поддерживает версии 1,0 и 1,1 для VersionOverrides, поэтому значение **xmlns** зависит от значения **xsi: Type** этого `<VersionOverrides>` элемента:
    - Если **xsi: Type** , `VersionOverridesV1_0`то **xmlns** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides`.
    - Если **xsi: Type** , `VersionOverridesV1_1`то **xmlns** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.

> [!NOTE]
> В настоящее время только Outlook 2016 или более поздней версии поддерживает схему VersionOverrides `VersionOverridesV1_1` 1.1 и тип.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  НЕТ   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](./resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.|
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.|
|  [Hosts](./hosts.md)                |  Да  |  Задает набор узлов Office. Дочерний элемент Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](./resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](./equivalentaddins.md)    |  Нет  | Задает встроенные надстройки (COM/XLL), эквивалентные веб-надстройке. Веб-надстройка не активируется, если установлена эквивалентная собственная встроенная надстройка.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](./webapplicationinfo.md)    |  Нет  | Задает сведения о регистрации надстройки с помощью надежных поставщиков маркеров, таких как Azure Active Directory 2.0. |

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
