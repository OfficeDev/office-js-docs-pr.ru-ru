---
title: Элемент Host в файле манифеста
description: Определяет тип приложения Office, в котором следует активировать надстройку.
ms.date: 11/05/2019
ms.localizationpriority: medium
ms.openlocfilehash: e262183f11361d02d1a710d1f442539a8e936bcd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154710"
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
| [Name](#name) | string | Обязательный | Имя типа Office клиентского приложения. |

### <a name="name"></a>Имя

Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

> [!IMPORTANT]
> Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint. В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.

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
|  [xsi:type](#xsitype)  |  Да  | Описывает приложение Office, в котором применяются эти параметры.|

### <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Да   |  Определяет параметры классического форм-фактора. |
|  [MobileFormFactor](mobileformfactor.md)    |  Нет   |  Определяет параметры мобильного форм-фактора. **Примечание:** Этот элемент поддерживается только в Outlook iOS и Android. |
|  [AllFormFactors](allformfactors.md)    |  Нет   |  Определяет параметры всех форм-факторов. Используется только пользовательскими функциями в Excel. |

### <a name="xsitype"></a>xsi:type

Элементы управления Office приложения (Word, Excel, PowerPoint, Outlook, OneNote), где применяются содержащиеся параметры. Поддерживаются такие значения:

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
