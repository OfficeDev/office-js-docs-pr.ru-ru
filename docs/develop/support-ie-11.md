---
title: Поддержка Internet Explorer 11
description: Узнайте, как поддерживать JavaScript для Internet Explorer 11 и ES5 в надстройке.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: aff6004af4ce28aea865cb34cd34e13e23fb549f
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810276"
---
# <a name="support-internet-explorer-11"></a>Поддержка Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему используется в надстройках Office**
>
> Некоторые сочетания платформ и версий Office, включая бессрочные версии Office 2019, по-прежнему используют элемент управления webview, который поставляется с Internet Explorer 11, для размещения надстроек, как описано в [статье Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md). Мы рекомендуем (но не требовать), чтобы вы продолжали поддерживать эти сочетания, по крайней мере в минимальном виде, предоставляя пользователям надстройки корректное сообщение о сбое при запуске надстройки в веб-представлении Internet Explorer. Помните о следующих дополнительных моментах:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) больше не тестирует надстройки в Office в Интернете использует Internet Explorer в качестве браузера.
> - AppSource по-прежнему тестирует сочетание версий платформы и *классических* версий Office, использующих Internet Explorer, однако выдает предупреждение только в том случае, если надстройка не поддерживает Internet Explorer. Надстройка не отклоняется AppSource.
> - [Средство Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.

Надстройки Office — это веб-приложения, которые отображаются в IFrame при выполнении на Office в Интернете. Надстройки Office отображаются с помощью встроенных элементов управления браузера при запуске в Office в Windows или Office на Компьютере Mac. Внедренные элементы управления браузера предоставляются операционной системой или браузером, установленным на компьютере пользователя.

Если вы планируете поддерживать более старые версии Windows и Office, надстройка должна работать во встраиваемом элементе управления браузера, основанном на Internet Explorer 11 (IE11). Сведения о том, какие сочетания Windows и Office используют элемент управления браузером на основе IE11, см. [в разделе Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает некоторые функции HTML5, такие как мультимедиа, запись и расположение. Если надстройка должна поддерживать Internet Explorer 11, необходимо либо разработать надстройку, чтобы избежать этих неподдерживаемых функций, либо надстройка должна определить, когда используется Internet Explorer, и предоставить альтернативный интерфейс, который не использует неподдерживаемые функции. Дополнительные сведения см. [в разделе Определение во время выполнения, если надстройка запущена в Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Поддержка последних версий JavaScript

Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней версии или TypeScript, у вас есть два варианта, как описано в этой статье. Вы также можете объединить эти два метода.

### <a name="use-a-transpiler"></a>Использование транспилера

Вы можете написать код в TypeScript или современном JavaScript, а затем транспилировать его во время сборки в ES5 JavaScript. Полученные файлы ES5 передаются в веб-приложение надстройки.

Есть два популярных транспилеров. Оба они могут работать с исходными файлами, которые являются TypeScript или JavaScript после ES5. Они также работают с файлами React (JSX и TSX).

- [Babel](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

Сведения об установке и настройке транспилировщика в проекте надстройки см. в документации по любой из них. Для автоматизации транспиляции рекомендуется использовать средство выполнения задач, например [Grunt](https://gruntjs.com/) или [WebPack](https://webpack.js.org/) . Пример надстройки, использующий tsc, см. [в статье Надстройка Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Пример, в котором используется babel, см. [в статье Надстройка автономного хранилища](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Если вы используете Visual Studio (не Visual Studio Code), tsc, вероятно, проще всего использовать. Вы можете установить поддержку для него с помощью пакета nuget. Дополнительные сведения см. [в статье JavaScript и TypeScript в Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Чтобы использовать babel с Visual Studio, создайте скрипт сборки или используйте обозреватель средства выполнения задач в Visual Studio с такими инструментами, как [Средство выполнения задач WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) или [Средство выполнения задач NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Использование полизаполнения

[Polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) — это Более ранняя версия JavaScript, которая дублирует функциональные возможности более поздних версий JavaScript. Polyfill работает с в браузерах, которые не поддерживают более поздние версии JavaScript. Например, строковый метод `startsWith` не входит в версию JavaScript ES5 и поэтому не будет выполняться в Internet Explorer 11. Существуют библиотеки polyfill, написанные на ES5, которые определяют и реализуют `startsWith` метод. Рекомендуется использовать библиотеку [polyfill core-js](https://github.com/zloirock/core-js) .

Чтобы использовать библиотеку polyfill, загрузите ее, как и любой другой файл или модуль JavaScript. Например, можно использовать `<script>` тег в HTML-файле домашней страницы надстройки (например `<script src="/js/core-js.js"></script>`, ) или `import` оператор в файле JavaScript (например, `import 'core-js';`). Когда обработчик JavaScript видит такой метод, как `startsWith`, он сначала посмотрит, есть ли метод с таким именем, встроенный в язык. Если это так, он вызовет собственный метод. Если метод не является встроенным и только если он не является встроенным, подсистема будет искать его во всех загруженных файлах. Таким образом, полизаполненные версии не используются в браузерах, поддерживающих собственную версию.

Импорт всей библиотеки core-js приведет к импорту всех функций core-js. Вы также можете импортировать только те polyfills, которые требуются надстройке Office. Инструкции о том, как это сделать, см. в разделе [Api CommonJS](https://github.com/zloirock/core-js#commonjs-api). Библиотека core-js содержит большинство необходимых полизаполнения. Существует несколько исключений, подробно описанных в разделе [Отсутствующие Polyfills](https://github.com/zloirock/core-js#missing-polyfills) документации core-js. Например, он не поддерживает `fetch`, но можно использовать [выборку](https://github.com/github/fetch) polyfill.

Пример надстройки, использующий core.js, см. [в статье Надстройка Word Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Определите, запущена ли надстройка в Internet Explorer во время выполнения.

Надстройка может обнаружить, запущена ли она в Internet Explorer, считывая свойство [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Это позволяет надстройке либо предоставлять альтернативный интерфейс, либо корректно завершать сбой. Ниже приведен пример. Обратите внимание, что Internet Explorer отправляет строку, начинающуюся с "Trident" в качестве значения userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade 
    //      either to perpetual Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> Обычно чтение свойства не рекомендуется `userAgent` . Убедитесь, что вы знакомы со статьей [Обнаружение браузера с помощью агента пользователя](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent), включая рекомендации и альтернативы для чтения `userAgent`. В частности, если вы используете вариант 1 в приведенном `else` выше предложении, рассмотрите возможность использования обнаружения признаков вместо тестирования для агента пользователя.
>
> По состоянию на 30 сентября 2021 г. текст в разделе [Какая часть агента пользователя содержит нужные сведения?](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) датируется до выпуска Internet Explorer 11. Это по-прежнему в целом точно, и *таблицы* в разделе английской версии статьи актуальны. Аналогичным образом текст и в большинстве случаев таблицы в версиях статьи, отличных от английского, устарели.

## <a name="test-an-add-in-on-internet-explorer"></a>Тестирование надстройки в Internet Explorer

См. статью [Тестирование Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Таблица совместимости ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Можно ли использовать... Поддержка таблиц для HTML5, CSS3 и т. д.](https://caniuse.com/)
