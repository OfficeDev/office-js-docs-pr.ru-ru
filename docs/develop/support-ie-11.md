---
title: Поддержка Internet Explorer 11
description: Узнайте, как поддерживать Internet Explorer 11 и Javascript ES5 в надстройки.
ms.date: 10/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5bb39235961fcb6ab37b211fe96d2c776de5a9ad
ms.sourcegitcommit: a37be80cf47a37c85b7f5cab216c160f4e905474
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/09/2021
ms.locfileid: "60250422"
---
# <a name="support-internet-explorer-11"></a>Поддержка Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer по-прежнему Office надстройки**
>
> Корпорация Майкрософт заканчивает поддержку Internet Explorer, но это не влияет на Office надстройки. Некоторые сочетания платформ и версий Office, включая версии с одновековой покупкой до Office 2019 г., будут по-прежнему использовать управление веб-просмотром, которое поставляется с Internet Explorer 11 для пользования надстройки, как это объясняется в браузерах, используемых [Office надстройки](../concepts/browsers-used-by-office-web-add-ins.md). Кроме того, поддержка этих комбинаций и, следовательно, internet Explorer по-прежнему требуется для надстройок, представленных [в AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Меняются *две* вещи:
>
> - Office в Интернете больше не открывается в Internet Explorer. Следовательно, AppSource больше не тестирует надстройки в Office в Интернете с помощью Internet Explorer в качестве браузера. Но AppSource по-прежнему тестирует комбинации  платформы и Office настольных версий, которые используют Internet Explorer.
> - Средство [Script Lab](../overview/explore-with-script-lab.md) больше не поддерживает Internet Explorer.

Office Надстройки — это веб-приложения, которые отображаются в IFrames при Office в Интернете. Office Надстройки отображаются с помощью встроенных элементов управления браузером при Office на Windows или Office mac. Встроенные элементы управления браузером поставляются операционной системой или браузером, установленным на компьютере пользователя.

Если вы планируете выставлять надстройку на рынок через AppSource или планируете поддерживать более старые версии Windows и Office, надстройка должна работать в встраиваемом контроле браузера, основанном на Internet Explorer 11 (IE11). Сведения о том, какие сочетания Windows и Office используют управление браузером на основе IE11, см. в браузерах, используемых Office [надстройки.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!IMPORTANT]
> Internet Explorer 11 не поддерживает некоторые функции HTML5, такие как мультимедиа, запись и расположение. Если надстройка должна поддерживать Internet Explorer 11, вы не можете использовать эти функции.

Internet Explorer 11 не поддерживает версии JavaScript позже ES5. Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней части или TypeScript, у вас есть два варианта, описанных в этой статье. Вы также можете объединить эти два метода.

## <a name="use-a-transpiler"></a>Использование транспилера

Код можно написать как в TypeScript, так и в современном JavaScript, а затем перенастроить его во время сборки в JavaScript ES5. В результате в веб-приложение надстройки загружаются файлы ES5.

Существует два популярных транспилера. Оба из них могут работать с исходными файлами, которые typeScript или post-ES5 JavaScript. Они также работают с React файлами (jsx и .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

Сведения об установке и настройке транспилера в проекте надстройки см. в документации для любого из них. Для автоматизации переливания рекомендуется использовать бегун задач, например [Grunt](https://gruntjs.com/) или [WebPack.](https://webpack.js.org/) Пример надстройки, использующей tsc, см. в Office надстройки [Microsoft Graph React.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React) Пример, использующий babel, см. в служба хранилища [надстройки.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin)

> [!NOTE]
> Если вы используете Visual Studio (не Visual Studio Code), tsc, вероятно, проще всего использовать. Вы можете установить поддержку для него с помощью пакета nuget. Дополнительные сведения см. в [javaScript и TypeScript в Visual Studio 2019 г.](/visualstudio/javascript/javascript-in-vs-2019) Чтобы использовать babel с Visual Studio, создайте сценарий сборки или используйте обозреватель задач runner в Visual Studio с помощью таких средств, как бегун задач [WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) или [NPM Task Runner.](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)

## <a name="use-a-polyfill"></a>Использование полифайла

[Полифильм](https://en.wikipedia.org/wiki/Polyfill_(programming)) — это JavaScript более ранней версии, который дублирует функции из более последних версий JavaScript. Полифильм работает с браузерами, которые не поддерживают более поздние версии JavaScript. Например, метод строки не был частью версии ES5 JavaScript, поэтому он не будет работать в `startsWith` Internet Explorer 11. Существуют библиотеки полифильмов, написанные в ES5, которые определяют и реализуют `startsWith` метод. Рекомендуется библиотека [полифильмов core-js.](https://github.com/zloirock/core-js)

Чтобы использовать библиотеку полифильмов, загрузите ее, как и любой другой файл JavaScript или модуль. Например, можно использовать тег в HTML-файле домашней страницы надстройки (например), или можно использовать заявление в `<script>` `<script src="/js/core-js.js"></script>` `import` файле JavaScript (например). `import 'core-js';` Когда двигатель JavaScript видит такой метод, он сначала будет искать, есть ли метод этого имени, встроенный `startsWith` в язык. Если есть, он будет вызывать родной метод. Если метод не встроен и только в том случае, если он не встроен, двигатель будет искать для него все загруженные файлы. Таким образом, полифулловая версия не используется в браузерах, поддерживаюх родную версию.

Импорт всей библиотеки core-js импортирует все функции core-js. Вы также можете импортировать только полифильмы, Office надстройки. Инструкции по этому поводу см. в [aPI CommonJS.](https://github.com/zloirock/core-js#commonjs-api) Библиотека core-js имеет большинство необходимых полифильмов. В разделе Missing [Polyfills](https://github.com/zloirock/core-js#missing-polyfills) документации core-js описано несколько исключений. Например, он не `fetch` поддерживается, но вы можете использовать [подбирать](https://github.com/github/fetch) полифильм.

Пример надстройки, использующей core.js, см. в примере [Word Add-in Angular2 StyleChecker.](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)

## <a name="testing-an-add-in-on-internet-explorer"></a>Тестирование надстройки в Internet Explorer

См. [тест Internet Explorer 11](../testing/ie-11-testing.md).

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Таблица совместимости ECMAScript 6](https://kangax.github.io/compat-table/es6/)
- [Могу ли я использовать... Таблицы поддержки для HTML5, CSS3 и т.д.](https://caniuse.com/)
