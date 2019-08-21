---
title: Элемент VersionOverrides в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: ce65cdced1b3cf885cee09732c2cda0081a53cfc
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477882"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Расположение схемы (`http://schemas.microsoft.com/office/mailappversionoverrides`, когда `xsi:type` — `VersionOverridesV1_0`, и `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, когда `xsi:type` — `VersionOverridesV1_1`).|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

> [!NOTE]
> В настоящее время только Outlook 2016 или более поздней версии поддерживает схему VersionOverrides `VersionOverridesV1_1` 1.1 и тип.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  НЕТ   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](./resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.|
| **екуивалентаддинс** | Нет | Задает обратную совместимость с эквивалентной надстройкой COM, XLL или и тем, и другими. |
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.|
|  [Hosts](./hosts.md)                |  Да  |  Задает набор узлов Office. Дочерний элемент Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](./resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [екуивалентаддинс](./equivalentaddins.md)    |  Нет  | Задает встроенные надстройки (COM/XLL), эквивалентные веб-надстройке. Веб-надстройка не активируется, если установлена эквивалентная собственная встроенная надстройка.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](./webapplicationinfo.md)    |  Нет  | Задает сведения о регистрации надстройки с помощью надежных поставщиков маркеров, таких как Azure Active Directory 2.0. |

### <a name="versionoverrides-example"></a>Пример VersionOverrides

Ниже приведен пример типичного `<VersionOverrides>` элемента, в том числе некоторые необязательные дочерние элементы, которые обычно используются.

```xml
<OfficeApp>
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
<OfficeApp>
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
