---
title: Элемент Host в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127571"
---
# <a name="host-element"></a>Элемент Host

Определяет тип приложения Office, в котором следует активировать надстройку.

> [!IMPORTANT] 
> Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node). Функциональность в обоих случаях одинакова.  

## <a name="basic-manifest"></a>Базовый манифест

Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.

### <a name="attributes"></a>Атрибуты

| Атрибут     | Тип   | Обязательный | Описание                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | Обязательный | Имя типа ведущего приложения Office. |

### <a name="name"></a>Имя
Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>Пример
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>Узел VersionOverrides
Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`. 

### <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Описывает приложение Office, к которому применяются эти параметры.|

### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Да   |  Определяет параметры классического форм-фактора. |
|  [MobileFormFactor](mobileformfactor.md)    |  Нет   |  Определяет параметры для мобильного конструктивного параметра. **Примечание:** Этот элемент поддерживается только в Outlook в iOS. |
|  [AllFormFactors](allformfactors.md)    |  Нет   |  Определяет параметры всех форм-факторов. Используется только пользовательскими функциями в Excel. |

### <a name="xsitype"></a>xsi:type

Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры. Допустимые значения:

- `Document` (Word)
- `MailHost` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Пример ведущего приложения 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
