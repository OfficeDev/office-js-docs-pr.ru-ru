---
title: Изучение API JavaScript для Office с помощью Script Lab
description: Используйте сценарий "Лаборатория" для изучения API Office JS и прототипов функций.
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
ms.openlocfilehash: fbefd205ac929579cea1120b8398a53146bca19c
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851518"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Изучение API JavaScript для Office с помощью Script Lab

[Надстройка "Лаборатория скриптов](https://appsource.microsoft.com/product/office/WA104380862)", доступная бесплатно из AppSource, позволяет изучать API JavaScript для Office при работе с программами Office, такими как Excel или Word. Script Lab — удобное средство для добавления в набор средств разработки в качестве прототипа и проверки функциональных возможностей, которые должны быть в надстройке.

## <a name="what-is-script-lab"></a>Что такое "Лаборатория скриптов"?

Script Lab — это средство для тех, кто хочет научиться разрабатывать надстройки Office с помощью API JavaScript для Office в Excel, Word или PowerPoint. Он предоставляет IntelliSense, чтобы вы могли видеть доступные и созданные на платформе Монако платформы, ту же платформу, которая используется в Visual Studio Code. С помощью сценария Lab вы можете получить доступ к библиотеке образцов, чтобы быстро испытать функции, или вы можете использовать пример в качестве отправной точки для собственного кода. Вы также можете воспользоваться лабораториями скриптов для предварительной версии API.

Звучит хорошо? Просмотрите этот видеоролик в виде одной минуты, чтобы увидеть Лаборатория сценариев в действии.

[![Предварительный просмотр видео, в котором показана Лаборатория скриптов, работающая в Excel, Word и PowerPoint.](../images/screenshot-wide-youtube.png 'Видеоролик о предварительном просмотре в лаборатории сценариев')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>Основные возможности

В разделе script Lab предусмотрен ряд функций, которые помогут вам изучить функциональные возможности API JavaScript для Office и прототипа надстройки.

### <a name="explore-samples"></a>Обзор примеров

Быстро приступите к работе со статьей встроенных примеров фрагментов, демонстрирующих выполнение задач с помощью API. Вы можете запустить примеры, чтобы сразу увидеть результат в области задач или документе, изучить примеры, чтобы узнать, как работает API, и даже использовать примеры для создания прототипа собственной надстройки.

![Примеры](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>Код и стиль

В дополнение к коду JavaScript или TypeScript, вызывающему API Office JS, каждый фрагмент также содержит HTML-разметку, определяющую содержимое области задач и CSS, определяющую внешний вид области задач. Вы можете настроить HTML-разметку и CSS, чтобы поэкспериментировать с размещением элементов и стилизацией при создании прототипа области задач для собственной надстройки.

> [!TIP]
> Чтобы вызывать API предварительного просмотра внутри фрагмента, вам потребуется обновить библиотеки фрагментов кода, чтобы использовать бета-версию CDN`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`() и определения `@types/office-js-preview`типов предварительного просмотра. Кроме того, некоторые API предварительной версии доступны только в том случае, если вы зарегистрировались в [программе предварительной оценки Office](https://products.office.com/office-insider) и у вас установлена сборка Office для участников.

### <a name="save-and-share-snippets"></a>Сохранение и совместное использование фрагментов

По умолчанию фрагменты кода, открываемые в лаборатории сценариев, будут сохранены в кэше браузера. Для окончательного сохранения фрагмента его можно экспортировать в [GitHub](https://gist.github.com). Создайте секретный объект, чтобы сохранить фрагмент исключительно для собственного использования, или создайте общедоступного пользователя, если вы планируете поделиться им с другими пользователями.

![Параметры общего доступа](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>Импорт фрагментов кода

Вы можете импортировать фрагмент в тестовый сценарий, указав URL-адрес общедоступного [GitHub](https://gist.github.com) , в котором ХРАНИТСЯ фрагмент ямл, или ВСТАВИВ полный ямл для фрагмента. Эта функция может быть полезна в тех случаях, когда кто-то другой предоставил доступ к своему фрагменту, опубликовав его в GitHub или предоставляя свой фрагмент кода ЯМЛ.

![Параметр "импортировать фрагмент"](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>Поддерживаемые клиенты

Лаборатория скриптов поддерживается для Excel, Word и PowerPoint на следующих клиентах.

- Office 2013 или более поздней версии в Windows
- Office 2016 или более поздней версии на компьютерах Mac
- Office в Интернете

## <a name="next-steps"></a>Дальнейшие действия

Чтобы использовать сценарий "Лаборатория" в Excel, Word или PowerPoint, установите [надстройку "Лаборатория скриптов](https://appsource.microsoft.com/product/office/WA104380862) " из AppSource. 

Вы можете развернуть учебную библиотеку в лаборатории сценариев, дополнив новые фрагменты кода в репозиторий GitHub для [Office – JS: Snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) .

Когда вы будете готовы создать свою первую надстройку Office, ознакомьтесь с кратким руководством для [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md)или [Project](../quickstarts/project-quickstart.md).

## <a name="see-also"></a>См. также

- [Получение лаборатории сценариев](https://appsource.microsoft.com/product/office/WA104380862)
- [Дополнительные сведения о лаборатории сценариев](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Регистрация в программе для разработки](https://developer.microsoft.com/office/dev-program)
- [Создание надстроек Office](../overview/office-add-ins-fundamentals.md)
