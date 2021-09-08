---
title: Элемент GetStarted в файле манифеста
description: Предоставляет сведения, используемые при установке надстройки в Word, Excel, PowerPoint и OneNote.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a637f3f9031d9f8e09d14f17f2095ca0647c4d50
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937338"
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
