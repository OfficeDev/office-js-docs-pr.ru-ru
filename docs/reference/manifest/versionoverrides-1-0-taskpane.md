---
title: VersionOverrides элемент 1.0 в файле манифеста для надстройки области задач
description: Справочная документация элемента VersionOverrides (области задач) для Office файлов манифеста надстройок (XML).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 266a20ea2b2d980007bd05411150f2f152b6c7c1
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042193"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-task-pane-add-in"></a>VersionOverrides элемент 1.0 в файле манифеста для надстройки области задач

Этот элемент содержит сведения для функций, которые не поддерживаются в базовом манифесте.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с обзором элемента [VersionOverrides,](versionoverrides.md)который содержит важную информацию о атрибутах и вариантах элемента.

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides:**

- Taskpane 1.0

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (требуется для Excel, PowerPoint и Word.)
- Некоторые детские элементы могут быть связаны с дополнительными наборами требований.

## <a name="child-elements"></a>Дочерние элементы

В следующей таблице используется только версия 1.0 элементов **VersionOverrides** и только надстройки области задач.

> [!NOTE]
> Только в iOS `<WebApplicationInfo>` поддерживается. Все остальные детские элементы **VersionOverrides** игнорируются.

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Описание](#description)    |  Нет   |  Описывает надстройку. |
|  [Requirements](requirements.md)  |  Нет   |  Указывает минимальные наборы требований, которые необходимо поддерживать для того, чтобы разметка в родительском документе `VersionOverrides` вступила в силу. Это всегда должно быть *более строгим,* чем элемент `Requirements` в базовой части манифеста.|
|  [Hosts](hosts.md)                |  Да  |  Указывает коллекцию Office приложений. Элемент Child Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|
|  [EquivalentAddins](equivalentaddins.md)    |  Нет  | Указывает родной (COM/XLL) надстройки, эквивалентные веб-надстройки. Веб-надстройка не активируется, если установлена эквивалентная родной надстройка.|
|  **VersionOverrides**    |  Нет  | В настоящее время не может быть в VersionOverrides 1.0 для надстройок taskpane. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Нет  | Указывает сведения о регистрации надстройки с защищенными эмитентами маркеров, такими как Azure Active Directory V2.0. |

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

```xml
<OfficeApp ... xsi:type="Taskpane">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
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
