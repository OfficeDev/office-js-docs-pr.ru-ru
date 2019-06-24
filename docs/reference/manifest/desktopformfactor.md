---
title: Элемент DesktopFormFactor в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d1f09203518a38f1568b13e6c1a9c70752697152
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128519"
---
# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Настольный конструктивный фактор включает Office в Интернете, Windows и Mac. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.

В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Да      | Определяет, где предоставляются функции надстройки. |
| [FunctionFile](functionfile.md)       | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [GetStarted](getstarted.md)           | Нет       | Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Нет | Определяет, доступна ли надстройка Outlook в сценариях делегирования, и имеет значение *false* по умолчанию.<br><br>**Важно!** поскольку доступ представителя для надстроек Outlook в настоящее время находится в предварительной версии, надстройки, использующие `SupportSharedFolders` этот элемент, не могут быть опубликованы в AppSource или развернуты с помощью централизованного развертывания. |

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
