---
title: Элемент DesktopFormFactor в файле манифеста
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: cddf76af01ec9f3016b28a3f7692aa6dfeb9bd60
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450557"
---
# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.

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
