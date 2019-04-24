---
title: Элемент GetStarted в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d9ebcba7881b388544eeb3e2c3028bff9bdcf9a6
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452083"
---
# <a name="getstarted-element"></a>Элемент GetStarted

Предоставляет сведения для выноски, которая отображается при установке надстройки в ведущих приложениях Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Дочерние элементы

| Элемент                       | Обязательный | Описание                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Да      | Определяет, где предоставляются функции надстройки.     |
| [Описание](#description)   | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Нет       | URL-адрес страницы с подробным описанием надстройки.   |

### <a name="title"></a>Название 

Обязательный. Заголовок в верхней части выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).

### <a name="description"></a>Описание

Обязательный. Описание и основной текст выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **LongStrings** в разделе [Resources](resources.md).

### <a name="learnmoreurl"></a>LearnMoreUrl

Обязательный. URL-адрес страницы, где пользователь может узнать больше о надстройке. Атрибут **resid** ссылается на допустимый идентификатор элемента **Urls** в разделе [Resources](resources.md).

> [!NOTE]
> В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint. Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен. 

## <a name="see-also"></a>См. также

В следующих примерах кода используется элемент **GetStarted**:

* [Веб-надстройка Excel для работы с форматированием таблиц и диаграмм](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [JavaScript SpecKit для надстроек Word](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
