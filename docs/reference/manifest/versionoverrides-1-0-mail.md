---
title: VersionOverrides элемент 1.0 в файле манифеста для надстройки почты
description: Справочная документация элемента VersionOverrides (почта) для Office файлов манифеста надстройок (XML).
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5288c085c94ff6fc8ab8fc31711c5c8fa142e946
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340675"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-mail-add-in"></a>VersionOverrides элемент 1.0 в файле манифеста для надстройки почты

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с обзором элемента [VersionOverrides](versionoverrides.md), который содержит важную информацию о атрибутах и вариантах элемента.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.0

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- Некоторые детские элементы могут быть связаны с дополнительными наборами требований.

## <a name="child-elements"></a>Дочерние элементы

В следующей таблице используется только версия 1.0 элементов **VersionOverrides** и только почтовые надстройки.

> [!NOTE]
> В iOS поддерживается только **WebApplicationInfo** . Все остальные детские элементы **VersionOverrides** игнорируются.

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Описание](#description)    |  Нет   |  Описывает надстройку. |
|  [Requirements](requirements.md)  |  Нет   |  Указывает минимальные наборы требований, которые необходимо поддерживать для того, чтобы разметка в родительских **VersionOverrides** вступила в силу. Это всегда должно быть *более ограничительным* , чем элемент **Requirements** в базовой части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Указывает коллекцию Office приложений. Дочерний элемент **Hosts** переопределяет элемент **Hosts** в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  **VersionOverrides**    |  Нет  | Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки с защищенными эмитентами маркеров, такими как Azure Active Directory V2.0. |

### <a name="description"></a>Описание

Описывает надстройку. Он переопределяет элемент **Description** в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Атрибут `resid` элемента **Description** может быть не более 32 `id` символов и должен соответствовать значению атрибута детского элемента **элемента ShortString** , содержаного в [элементе Resources](resources.md) . 

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) , когда родительский **VersionOverrides** — это тип Taskpane 1.0.
- [Почтовый ящик 1.3,](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) когда родительский **VersionOverrides** — это тип Почта 1.0.
- [Почтовый ящик 1.5,](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) когда родительский **VersionOverrides** — это тип Почта 1.1.

## <a name="example"></a>Пример

Ниже приведен простой пример. Более сложные примеры см. в манифестах для примерных надстройок в Office [примерах кода надстройки](https://github.com/OfficeDev/PnP-OfficeAddins).

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

Манифест может реализовать несколько версий элемента **VersionOverrides** , которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы дополнительно поддерживать новые функции в более новой схеме, поддерживая старых клиентов, которые не поддерживают новые функции.

Для реализации нескольких версий элемент **VersionOverrides** `VersionOverrides` для более новой версии должен быть ребенком элемента для более старой версии. Элемент **child VersionOverrides** не наследует значения от родительского элемента.

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
