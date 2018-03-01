---
title: Обзор надстроек Word
description: ''
ms.date: 01/23/2018
---


# <a name="word-add-ins-overview"></a>Обзор надстроек Word

Хотите создать решение для автоматического составления документов или привязки и доступа к данным в документе Word из других источников? Чтобы расширить возможности клиентов Word на компьютере с Windows, Mac или в облаке, используйте платформу надстроек Office, которая включает API JavaScript для Word и API JavaScript для Office.

На [платформе надстроек Office](../overview/office-add-ins.md) можно разрабатывать не только надстройки Word. Используя команды надстроек, вы можете расширять интерфейс Word и запускать области задач, которые выполняют сценарий JavaScript, взаимодействующий с содержимым документа. Любой код, который работает в браузере, будет работать в надстройке Word. Надстройки, взаимодействующие с содержимым документа Word, создают запросы на совершение действий с объектами Word и синхронизацию состояния этих объектов. 

> [!NOTE]
> Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource, она должна соответствовать [политикам проверки AppSource](https://docs.microsoft.com/ru-ru/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/ru-ru/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и [статье о доступности надстроек Office в ведущих приложениях](../overview/office-add-in-availability.md).

Ниже показан пример надстройки Word, работающей в области задач.

*Рис. 1. Надстройка, работающая в области задач Word*

![Надстройка, работающая в области задач Word](../images/word-add-in-show-host-client.png)

Надстройка Word может (1) отправлять запросы в документ Word и (2) обновлять, удалять или перемещать абзац, используя JavaScript для доступа к объекту paragraph. Например, в приведенном ниже коде показано, как добавить в абзац новое предложение.

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

Для размещения надстройки Word можно использовать любой веб-сервер, в частности ASP.NET, NodeJS и Python. Используйте любимую клиентскую платформу — Ember, Backbone, Angular, React —для разработки своего решения; или продолжайте работу с VanillaJS. Для [аутентификации](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) и размещения приложения можно использовать Azure.

API JavaScript для Word предоставляют приложению доступ к объектам и метаданным документа Word. С помощью этих API можно создавать надстройки, предназначенные для:

* Word 2013 для Windows
* Word 2016 для Windows
* Word Online
* Word 2016 для Mac
* Word для iOS

Написанные вами надстройки будут работать во всех версиях Word на различных платформах. Дополнительные сведения см. в статье [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).

## <a name="javascript-apis-for-word"></a>API JavaScript для Word

Для взаимодействия с объектами и метаданными в документе Word можно использовать два набора API JavaScript. Первый — [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word) — появился в Office 2013. Это общий API — многие объекты можно использовать в надстройках, размещенных в двух или более клиентах Office. В этом API широко используются обратные вызовы. 

Второй — [API JavaScript для Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview). Это строго типизированная объектная модель, с помощью которой можно создавать надстройки Word, предназначенные для Word 2016 для Mac и Windows. Эта объектная модель использует обещания и предоставляет доступ к объектам Word, в частности [Body](https://dev.office.com/reference/add-ins/word/body), [ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol), [InlinePicture](https://dev.office.com/reference/add-ins/word/inlinepicture) и [Paragraph](https://dev.office.com/reference/add-ins/word/paragraph). API JavaScript для Word включает определения TypeScript и файлы vsdoc, чтобы вы могли получать подсказки кода в своей интегрированной среде разработки.

В настоящее время все клиенты Word поддерживают общий API JavaScript для Office, а большинство из них поддерживают и API JavaScript для Word. Дополнительные сведения о поддерживаемых клиентах см. в [справочнике по API](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).

Рекомендуем начать с API JavaScript для Word, так как с объектной моделью проще работать. Используйте API JavaScript для Word, если вам нужно:

* получить доступ к объектам в документе Word.

Используйте общий API JavaScript для Office, если вам нужно:

* создать надстройки для Word 2013;
* выполнить начальные действия для приложения;
* проверить поддерживаемый набор требований;
* получить доступ к метаданным документа, его параметрам и сведениям о среде;
* создать привязку к разделам документа и записать события;
* использовать пользовательские XML-части;
* открыть диалоговое окно.

## <a name="next-steps"></a>Дальнейшие действия

Готовы [создать свою первую надстройку Word](word-add-ins.md)? Вы также можете воспользоваться нашим интерактивным [руководством по началу работы](http://dev.office.com/getting-started/addins?product=Word). Используйте [манифест надстройки](../develop/add-in-manifests.md), чтобы указать ведущее приложение, имя, разрешения и другие сведения.

Чтобы узнать больше о том, как создать качественную и привлекательную надстройку Word, см. [руководство по разработке](../design/add-in-design.md) и [рекомендации](../concepts/add-in-development-best-practices.md).

После разработки надстройку можно [опубликовать](../publish/publish.md) в сетевой папке, каталоге приложений или AppSource.

## <a name="whats-coming-up-for-word-add-ins"></a>Над чем мы работаем?

Мы публикуем новые API для надстроек Word на странице [Открытые спецификации API](https://dev.office.com/reference/add-ins/openspec), чтобы вы могли делиться своим мнением. Узнайте, над какими функциями API JavaScript для Word мы работаем, и поделитесь своим мнением о проектируемых спецификациях.

## <a name="see-also"></a>См. также

* [Обзор платформы надстроек Office](../overview/office-add-ins.md)
* [Справочные материалы по API JavaScript для Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

