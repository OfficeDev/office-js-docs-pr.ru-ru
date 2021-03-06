Для [некоторых версий Office и Windows](../concepts/browsers-used-by-office-web-add-ins.md)javaScript двигатель, в котором запускают надстройки, предоставляется Internet Explorer. Двигатель Internet Explorer не поддерживает версии JavaScript позже ES5. Это означает, что без специальной обработки файлы JavaScript, которые служит ваша надстройка, не могут использовать синтаксис, типы или методы, которые были добавлены в язык после ES5. Это не означает, что вы должны *писать* в синтаксисе ES5. У вас есть два других варианта:

- Напишите код [в ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) (также называемый ES6) или позже JavaScript, или в TypeScript, а затем скомпилировать код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).
- Напишите в ECMAScript 2015 или более [](https://en.wikipedia.org/wiki/Polyfill_(programming)) поздний JavaScript, а также загрузите библиотеку полифильмов, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать код.

Дополнительные сведения об этих параметрах см. в [меню Support Internet Explorer 11.](../develop/support-ie-11.md)
