---
title: Поддержка Internet Explorer 11
description: Узнайте, как поддерживать JavaScript в Internet Explorer 11 и ES5 в надстройке.
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 70fea604c17525836857b7cff4c8670da757f2a6
ms.sourcegitcommit: fd04b41f513dbe9e623c212c1cbd877ae2285da0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2022
ms.locfileid: "65313179"
---
# <a name="support-internet-explorer-11"></a>Поддержка Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему используется в Office надстройки**
>
> В некоторых сочетаниях платформ и версий Office, включая версии с однофакторной покупкой до Office 2019, по-прежнему используется элемент управления webview, который поставляется с Internet Explorer 11 для размещения надстроек, как описано в [браузерах](../concepts/browsers-used-by-office-web-add-ins.md), используемых надстройки Office. Рекомендуется (но не обязательно) продолжать поддерживать эти сочетания, по крайней мере минимально, предоставляя пользователям надстройки корректное сообщение об ошибке при запуске надстройки в веб-представлении Internet Explorer. Учитывайте следующие дополнительные моменты:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) больше не тестирует надстройки в Office в Интернете в качестве браузера.
> - AppSource по-прежнему проверяет комбинации платформы и классических версий Office, использующих Internet Explorer, однако выдано предупреждение только в том случае, если надстройка не поддерживает Internet Explorer; AppSource не отклоняет надстройку. 
> - Средство [Script Lab больше](../overview/explore-with-script-lab.md) не поддерживает Internet Explorer.

Office надстройки — это веб-приложения, которые отображаются в IFrame при Office в Интернете. Office надстройки отображаются с помощью встроенных элементов управления браузера при запуске Office на Windows или Office на компьютере Mac. Встроенные элементы управления браузером предоставляются операционной системой или браузером, установленным на компьютере пользователя.

Если вы планируете поддерживать более старые версии Windows и Office, надстройка должна работать в элементе управления встраиваемым браузером, основанным на Internet Explorer 11 (IE11). Сведения о том, какие сочетания Windows и Office браузера на основе IE11, см. в разделе "Браузеры", используемые Office [надстройки](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает некоторые функции HTML5, такие как мультимедиа, запись и расположение. Если надстройка должна поддерживать Internet Explorer 11, необходимо либо разработать надстройку, чтобы избежать этих неподдерживаемых функций, либо надстройка должна определить, когда используется Internet Explorer, и предоставить альтернативный интерфейс, который не использует неподдерживаемые функции. Дополнительные сведения см. в разделе "Определение во время выполнения, выполняется ли надстройка [в Internet Explorer"](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## <a name="support-for-recent-versions-of-javascript"></a>Поддержка последних версий JavaScript

Internet Explorer 11 не поддерживает версии JavaScript позже ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней версии или TypeScript, у вас есть два варианта, как описано в этой статье. Вы также можете объединить эти два метода.

### <a name="use-a-transpiler"></a>Использование транспиллера

Вы можете написать код на TypeScript или современном JavaScript, а затем выполнить его транспилирование во время сборки в ES5 JavaScript. Полученные файлы ES5 передаются в веб-приложение надстройки.

Существует два популярных транспиллера. Оба они могут работать с исходными файлами, которые являются TypeScript или JavaScript после ES5. Они также работают с React файлов (JSX и TSX).

- [Babel](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

Сведения об установке и настройке транспиллера в проекте надстройки см. в документации по ним. Для автоматизации транспилирования рекомендуется использовать средство выполнения задач, например [Grunt](https://gruntjs.com/) или [WebPack](https://webpack.js.org/) . Пример надстройки, использующей TSC, см. в Office [microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). Пример, в котором используется надстройка, [служба хранилища автономном режиме](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> Если вы используете Visual Studio (не Visual Studio Code), tsc, вероятно, проще всего использовать. Вы можете установить для него поддержку с помощью пакета NuGet. Дополнительные сведения см. в [разделах JavaScript и TypeScript Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). Чтобы использовать приложение Visual Studio, создайте скрипт сборки или используйте обозреватель средств выполнения Visual Studio с такими средствами, как средство выполнения задач [WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) или [средство выполнения задач NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### <a name="use-a-polyfill"></a>Использование полизаполнения

[Polyfill —](https://en.wikipedia.org/wiki/Polyfill_(programming)) это JavaScript более ранней версии, который дублирует функции из более поздних версий JavaScript. Полизаполнение работает в браузерах, которые не поддерживают более поздние версии JavaScript. Например, строковый метод `startsWith` не был частью версии ES5 JavaScript, поэтому он не будет выполняться в Internet Explorer 11. Существуют библиотеки полизаполнения, написанные в ES5, которые определяют и реализуют `startsWith` метод. Мы рекомендуем [использовать библиотеку polyfill core-js](https://github.com/zloirock/core-js) .

Чтобы использовать библиотеку полизаполнения, загрузите ее, как и любой другой файл или модуль JavaScript. Например, можно `<script>` использовать тег в HTML-файле домашней страницы надстройки или `<script src="/js/core-js.js"></script>``import` оператор в файле JavaScript (например, `import 'core-js';`). Когда модуль JavaScript `startsWith`видит такой метод, он сначала будет искать, есть ли метод с таким именем, встроенным в язык. Если он есть, он будет вызывать собственный метод. Если и только в том случае, если метод не является встроенным, подсистема будет искать все загруженные файлы для него. Таким образом, версия с полизаполнением не используется в браузерах, поддерживающих собственную версию.

При импорте всей библиотеки core-js будут импортированы все функции Core-js. Вы также можете импортировать только полизаполнения, необходимые Office надстройке. Инструкции о том, как это сделать, см. в API [CommonJS](https://github.com/zloirock/core-js#commonjs-api). Библиотека core-js содержит большую часть необходимых полизаполнения. В разделе "Отсутствующие полизаполнения[](https://github.com/zloirock/core-js#missing-polyfills)" документации core-js описано несколько исключений. Например, он не поддерживается `fetch`, но можно использовать полизаполнение [выборки](https://github.com/github/fetch) .

Пример надстройки, использующей core.js, см. в разделе [Надстройка Word Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>Определение во время выполнения, выполняется ли надстройка в Internet Explorer

Чтобы узнать, работает ли надстройка в Internet Explorer, прочитайте свойство [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) . Это позволяет надстройке предоставлять альтернативный интерфейс или корректно завершить работу с ошибкой. Ниже приведен пример. Обратите внимание, что Internet Explorer отправляет строку, начинающееся с Trident в качестве значения userAgent.

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
> Обычно чтение свойства не рекомендуется `userAgent` . Убедитесь, что вы знакомы со статьей об обнаружении [браузера](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent) с помощью агента пользователя, включая рекомендации и альтернативы чтению `userAgent`. В частности, если вы используете вариант 1 `else` в приведенном выше предложении, рассмотрите возможность использования функции обнаружения, а не тестирования для агента пользователя.
>
> По данным на 30 сентября 2021 г. текст в разделе "Какая часть агента пользователя содержит информацию, которую вы ищете [?](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) даты до выпуска Internet Explorer 11". Она по-прежнему является общедоступной, и таблицы в разделе статьи на английском языке актуальны. Аналогичным образом, текст и в большинстве случаев таблицы в версиях статьи, отличных от английского, устарели.

## <a name="test-an-add-in-on-internet-explorer"></a>Тестирование надстройки в Internet Explorer

См. [сведения о тестировании Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Таблица совместимости ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Можно ли использовать... Таблицы поддержки для HTML5, CSS3 и т. д.](https://caniuse.com/)
