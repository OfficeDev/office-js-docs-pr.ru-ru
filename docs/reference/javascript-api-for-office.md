---
title: API JavaScript для Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: c8b33bbf9d0107786c0272410c59b1a3fe998cba
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870620"
---
# <a name="javascript-api-for-office"></a>API JavaScript для Office

API JavaScript для Office позволяет создавать веб-приложения, взаимодействующие с объектными моделями в ведущих приложениях Office. Ваше приложение будет ссылаться на библиотеку office.js, которая загружает скрипты. Библиотека office.js загружает объектные модели, подходящие для приложения Office, в котором запущена надстройка. Вы можете использовать следующие объектные модели JavaScript:

- **Общие интерфейсы API**, представленные в **Office 2013**. Модель загружается для **всех ведущих приложений Office** и подключает надстройку к клиентскому приложению Office. Объектная модель содержит API, предназначенные для определенных клиентов Office, а также API, которые подходят для нескольких ведущих клиентских приложений Office. Все это содержимое находится в разделе **Общий API**. Эта объектная модель использует обратные вызовы. 

  **Outlook** также использует синтаксис общих API. Все, к чему относится псевдоним Office, содержит объекты, которые можно использовать для написания скриптов надстроек Office, взаимодействующих с содержимым документов, листов, презентаций, почтовых элементов и проектов Office. Нужно использовать общие API, если надстройка предназначена для Office 2013 и более поздних версий. Эта объектная модель использует обратные вызовы.

- **API для конкретных ведущих приложений**, представленные в **Office 2016**. Эта объектная модель предусматривает использование строго типизированных объектов, предназначенных для конкретных ведущих приложений. Эти объекты соответствуют уже известным объектам, отображающимся при использовании клиентов Office, и будут применяться впредь в API JavaScript для Office. API, предназначенные для конкретных ведущих приложений, на данный момент включают API JavaScript для Word и API JavaScript для Excel.

## <a name="supported-host-applications"></a>Поддерживаемые ведущие приложения

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Общий API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint и Project](requirement-sets/powerpoint-and-project-note.md) поддерживают надстройки, созданные с помощью API JavaScript. Тем не менее в настоящее время у них нет API для конкретных ведущих приложений. Взаимодействие с этими ведущими приложениями происходит через общие API.

Дополнительные сведения о [поддерживаемых ведущих приложениях и других требованиях](../concepts/requirements-for-running-office-add-ins.md).

## <a name="open-api-specifications"></a>Открытые спецификации API

Мы публикуем новые API для надстроек Office на странице [Открытые спецификации API](openspec.md), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями мы работаем, и поделитесь своим мнением о создаваемых спецификациях.

## <a name="see-also"></a>См. также

- [Справочник по API JavaScript для Office](/javascript/api/overview/office)
