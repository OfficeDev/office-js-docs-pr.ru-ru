---
title: Элемент DesktopFormFactor в файле манифеста
description: Указывает параметры для надстройки классического форм-фактора.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 52c9a029e3f43e9b7d5416455eb99ef3de4dae7a
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138732"
---
# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Форм-фактор рабочего стола включает Office в Интернете, Windows и Mac. Он содержит все сведения о надстройки для форм-фактора рабочего стола, за исключением **узла Resources.**

Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько **элементов ExtensionPoint.** Дополнительные сведения см. в [элементе FunctionFile и](functionfile.md) [элементе ExtensionPoint.](extensionpoint.md)

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

- Области задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Да      | Определяет, где предоставляются функции надстройки. |
| [FunctionFile](functionfile.md)       | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [GetStarted](getstarted.md)           | Нет       | Определяет вызов, который появляется при установке надстройки в Word, Excel или PowerPoint. При опущении в вызываемом вызове используются значения элементов [DisplayName](displayname.md) и [Description.](description.md) |
| [SupportsSharedFolders](supportssharedfolders.md) | Нет | Определяет, доступна ли надстройка Outlook в общих почтовых ящиках (в настоящее время в предварительном просмотре) и общих папках (т. е. в сценариях делегирования доступа). Значение false *по* умолчанию. |

## <a name="desktopformfactor-example"></a>Пример DesktopFormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
