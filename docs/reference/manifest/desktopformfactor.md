---
title: Элемент DesktopFormFactor в файле манифеста
description: Указывает параметры для надстройки классического форм-фактора.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 18828e6b61a45ae2dc1528b3f7a54e664af09519
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292316"
---
# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Настольный конструктивный фактор включает Office в Интернете, Windows и Mac. Он содержит все сведения о надстройках для настольных форм, за исключением узла **Resources** .

Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** . Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Да      | Определяет, где предоставляются функции надстройки. |
| [FunctionFile](functionfile.md)       | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [GetStarted](getstarted.md)           | Нет       | Определяет выноску, которая отображается при установке надстройки в Word, Excel или PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Нет | Определяет, доступна ли надстройка Outlook в сценариях делегирования. По умолчанию задано значение *false* . |

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
