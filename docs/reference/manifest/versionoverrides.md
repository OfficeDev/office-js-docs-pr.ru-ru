---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация по элементу VersionOverrides для XML-файлов манифеста надстройки Office.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 772eaa416909d24f8035ed3e1445d1e4f06a244e
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771307"
---
# <a name="versionoverrides-element"></a>Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Пространство имен схемы VersionOverrides. Допустимые значения зависят от значения `<VersionOverrides>` **xsi:type** этого элемента и **значения xsi:type** родительского `<OfficeApp>` элемента. См. [значения пространства имен ниже.](#namespace-values)|
|  **xsi:type**  |  Да  | Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Значения пространства имен

Ниже перечислены необходимые значения **значения xmlns** в зависимости от **значения xsi:type** родительского `<OfficeApp>` элемента.

- **TaskPaneApp** поддерживает только версию 1.0 VersionOverrides, а **xmlns** должны быть `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** поддерживает только версию 1.0 VersionOverrides, а **xmlns** должны быть `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** поддерживает версии 1.0 и 1.1 VersionOverrides, поэтому значение **XMLNS** зависит от значения `<VersionOverrides>` **xsi:type** этого элемента:
    - Если **xsi:type** , `VersionOverridesV1_0` **XMLNS** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Если **xsi:type** , `VersionOverridesV1_1` **XMLNS** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> В настоящее время только Outlook 2016 или более поздней версии поддерживает схему VersionOverrides версии 1.1 и `VersionOverridesV1_1` тип.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  Нет   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Атрибут элемента Description не может быть больше 32 символов и имеет значение атрибута элемента, который `resid`  `id` содержит `String` текст.|
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Указывает коллекцию приложений Office. Элемент child Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](equivalentaddins.md)    |  Нет  | Указывает нативные надстройки (COM/XLL), эквивалентные веб-надстройки. Веб-надстройка не активируется, если установлена эквивалентная нативная надстройка.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки у надежных токенов, таких как Azure Active Directory 2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Нет  |  Указывает коллекцию расширенных разрешений.<br><br>**Важно!** Так как API [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) в настоящее время находится в предварительной версии, надстройки, которые используют этот элемент, не могут быть опубликованы в AppSource или развернуты через централизованное `ExtendedPermissions` развертывание. |

### <a name="versionoverrides-example"></a>Пример VersionOverrides

Ниже приводится пример типичного элемента, включая некоторые из них, которые не требуются, `<VersionOverrides>` но обычно используются.

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
