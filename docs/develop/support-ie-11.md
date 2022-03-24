---
title: Поддержка Internet Explorer 11
description: Узнайте, как поддерживать Internet Explorer 11 и Javascript ES5 в надстройки.
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: d2a504a6e030e6cf8d06c766cb500d6c11710ea9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744232"
---
# <a name="support-internet-explorer-11"></a>Поддержка Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему Office надстройки**
>
> Корпорация Майкрософт заканчивает поддержку Internet Explorer, но это не Office надстройки. Некоторые комбинации платформ и Office версий, включая версии с одновековой покупкой до Office 2019 г., будут по-прежнему использовать управление веб-просмотром, которое поставляется с Internet Explorer 11 для пользования надстройки, как это объясняется в [браузерах](../concepts/browsers-used-by-office-web-add-ins.md), используемых Office надстройки. Кроме того, поддержка этих комбинаций и, следовательно, internet Explorer по-прежнему требуется для надстройок, представленных [в AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Меняются *две* вещи:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, AppSource больше не тестирует надстройки в Office в Интернете с помощью Internet Explorer в качестве браузера. Но AppSource по-прежнему тестирует комбинации платформы и Office настольных версий, которые используют Internet Explorer. 
> - Средство [Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.

Office надстройки — это веб-приложения, отображаемые в IFrames при Office в Интернете. Office надстройки отображаются с помощью встроенных элементов управления браузером при Office на Windows или Office на Mac. Встроенные элементы управления браузером поставляются операционной системой или браузером, установленным на компьютере пользователя.

Если вы планируете выставлять надстройку на рынок через AppSource или планируете поддерживать более старые версии Windows и Office, ваша надстройка должна работать в встраиваемом контроле браузера, основанном на Internet Explorer 11 (IE11). Сведения о том, какие сочетания Windows и Office используют управление браузером на основе IE11, см. в браузерах, используемых Office [надстройки](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает некоторые функции HTML5, такие как мультимедиа, запись и расположение. Если надстройка должна поддерживать Internet Explorer 11, необходимо либо разработать надстройку, чтобы избежать этих неподдержки, либо надстройка должна определить, когда используется Internet Explorer, и предоставить альтернативный опыт, который не использует неподдержку. Дополнительные сведения см. в [добавлении Определение](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer) времени запуска надстройки в Internet Explorer.

## <a name="support-for-recent-versions-of-javascript"></a>Поддержка последних версий JavaScript

Internet Explorer 11 не поддерживает версии JavaScript позже ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней части или TypeScript, у вас есть два варианта, описанных в этой статье. Вы также можете объединить эти два метода.

### <a name="use-a-transpiler"></a>Использование транспилера

Код можно написать как в TypeScript, так и в современном JavaScript, а затем перенастроить его во время сборки в JavaScript ES5. В результате в веб-приложение надстройки загружаются файлы ES5.

Существует два популярных транспилера. Оба из них могут работать с исходными файлами, которые typeScript или post-ES5 JavaScript. Они также работают с React файлами (jsx и .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

Сведения об установке и настройке транспилера в проекте надстройки см. в документации для любого из них. Для автоматизации переливания рекомендуется использовать бегун задач, например [Grunt](https://gruntjs.com/) или [WebPack](https://webpack.js.org/) . Пример надстройки, использующей tsc, см. в Office [надстройки Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Пример, использующий babel, см. в служба хранилища [надстройки](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Если вы используете Visual Studio (не Visual Studio Code), tsc, вероятно, проще всего использовать. Вы можете установить поддержку для него с помощью пакета nuget. Дополнительные сведения см. в [javaScript и TypeScript в Visual Studio 2019 г](/visualstudio/javascript/javascript-in-vs-2019). Чтобы использовать babel с Visual Studio, создайте сценарий сборки или используйте обозреватель бегуна задач в Visual Studio с помощью таких средств, как [бегун задач WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) или [раннер задач NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Использование полифайла

[Полифильм](https://en.wikipedia.org/wiki/Polyfill_(programming)) — это JavaScript более ранней версии, который дублирует функции из более последних версий JavaScript. Полифильм работает с браузерами, которые не поддерживают более поздние версии JavaScript. Например, метод строки `startsWith` не был частью версии ES5 JavaScript, поэтому он не будет работать в Internet Explorer 11. Существуют библиотеки полифильмов, написанные в ES5, которые определяют и реализуют `startsWith` метод. Рекомендуется библиотека [полифильмов core-js](https://github.com/zloirock/core-js) .

Чтобы использовать библиотеку полифильмов, загрузите ее, как и любой другой файл JavaScript или модуль. Например, можно `<script>` использовать тег в HTML-файле домашней страницы надстройки (`<script src="/js/core-js.js"></script>`например), `import` или можно использовать заявление в файле JavaScript (например). `import 'core-js';` Когда двигатель JavaScript `startsWith`видит такой метод, он сначала будет искать, есть ли метод этого имени, встроенный в язык. Если есть, он будет вызывать родной метод. Если метод не встроен и только в том случае, если он не встроен, двигатель будет искать для него все загруженные файлы. Таким образом, полифулловая версия не используется в браузерах, поддерживаюх родную версию.

Импорт всей библиотеки core-js импортирует все функции core-js. Вы также можете импортировать только полифильмы, Office надстройки. Инструкции по этому поводу см. в [API CommonJS](https://github.com/zloirock/core-js#commonjs-api). Библиотека core-js имеет большинство необходимых полифильмов. В разделе Missing [Polyfills](https://github.com/zloirock/core-js#missing-polyfills) документации core-js описано несколько исключений. Например, он не поддерживается `fetch`, но вы можете использовать [подбирать](https://github.com/github/fetch) полифильм.

Пример надстройки, использующей core.js, см. в примере [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Определите во время запуска, запущена ли надстройка в Internet Explorer

Ваша надстройка может узнать, работает ли она в Internet Explorer, прочитав свойство [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Это позволяет надстройки либо предоставить альтернативный опыт, либо изящно сбой. Ниже приведен пример. Обратите внимание, что Internet Explorer отправляет строку начиная с "Trident" в качестве значения userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade to 
    //      either one-time purchase Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> Обычно чтение свойства не является хорошей `userAgent` практикой. Убедитесь, что вы знакомы со статьей [, обнаружение браузера](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent) с помощью агента пользователя, в том числе рекомендации и альтернативы чтению `userAgent`. В частности, если вы принимаете вариант 1 в `else` вышеуказанном пункте, рассмотрите возможность обнаружения функций вместо тестирования для агента пользователя.
>
> По данным на 30 сентября 2021 г. текст в разделе Какая часть агента пользователя содержит сведения, которые вы ищете [?](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) датируется до выпуска Internet Explorer 11. Она по-прежнему в целом точна, и таблицы в разделе английская версия статьи устарели. Кроме того, текст, а в большинстве случаев таблицы, в не-английских версиях статьи устарели.

## <a name="test-an-add-in-on-internet-explorer"></a>Тестирование надстройки в Internet Explorer

См [. тест Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Таблица совместимости ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Могу ли я использовать... Таблицы поддержки для HTML5, CSS3 и т.д.](https://caniuse.com/)
