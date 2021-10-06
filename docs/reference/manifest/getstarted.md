---
title: Элемент GetStarted в файле манифеста
description: Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1630b50824cda18ca92ef6b34b0105acf9a4ca9c
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138752"
---
# <a name="getstarted-element"></a>Элемент GetStarted

Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md). Если элемент **GetStarted** опущен, в вызываемом вызове используются значения элементов [DisplayName](displayname.md) и [Description.](description.md)

**Тип надстройки:** надстройки области задач

**Допустимо только в этих схемах VersionOverrides:**

- Области задач 1.0

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [AppCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

## <a name="child-elements"></a>Дочерние элементы

| Элемент                       | Обязательный | Описание                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Да      | Определяет, где предоставляются функции надстройки.     |
| [Описание](#description)   | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Да       | URL-адрес страницы с подробным описанием надстройки.   |

### <a name="title"></a>Title 

Обязательный. Заголовок в верхней части выноски. Атрибут **resid** ссылается на действительный ID в **элементе ShortStrings** в разделе [Ресурсы](resources.md) и может быть не более 32 символов.

### <a name="description"></a>Описание

Обязательный. Описание и основной текст выноски. Атрибут **resid** ссылается на допустимый ID в **элементе LongStrings** в разделе [Ресурсы](resources.md) и может быть не более 32 символов.

### <a name="learnmoreurl"></a>LearnMoreUrl

Обязательный. URL-адрес страницы, где пользователь может узнать больше о надстройке. Атрибут **resid** ссылается на допустимый ID в **элементе Urls** в разделе [Ресурсы](resources.md) и может быть не более 32 символов.

> [!NOTE]
> В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint. Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен. 

## <a name="see-also"></a>См. также

В следующих примерах кода используется **элемент GetStarted.**

* [Веб-надстройка Excel для работы с форматированием таблиц и диаграмм](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [JavaScript SpecKit для надстроек Word](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
