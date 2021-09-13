---
title: Элемент DesktopFormFactor в файле манифеста
description: Указывает параметры для надстройки классического форм-фактора.
ms.date: 06/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: f89dff5626867258c8df93d5f047e3d08103e71b
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151623"
---
# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Форм-фактор рабочего стола включает Office в Интернете, Windows и Mac. Он содержит все сведения о надстройки для форм-фактора рабочего стола, за исключением **узла Resources.**

Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько **элементов ExtensionPoint.** Дополнительные сведения см. в [элементе FunctionFile и](functionfile.md) [элементе ExtensionPoint.](extensionpoint.md)

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Да      | Определяет, где предоставляются функции надстройки. |
| [FunctionFile](functionfile.md)       | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [GetStarted](getstarted.md)           | Нет       | Определяет вызов, который появляется при установке надстройки в Word, Excel или PowerPoint. |
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
