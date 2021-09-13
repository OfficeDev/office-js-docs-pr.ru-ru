---
title: Элемент GetStarted в файле манифеста
description: Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 355b72d4130f3a220e6a1257af51e371665d3cc3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154793"
---
# <a name="getstarted-element"></a>Элемент GetStarted

Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).

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
