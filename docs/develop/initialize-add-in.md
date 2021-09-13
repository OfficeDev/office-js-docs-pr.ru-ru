---
title: Инициализация надстройки Office
description: Узнайте, как инициализировать Office надстройки.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7879edce59aa3915b2e5a0d8c1b5e1c2c5a9fd30
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151014"
---
# <a name="initialize-your-office-add-in"></a>Инициализация надстройки Office

Надстройки Office часто поддерживают логику запуска для выполнения следующих действий:

- Убедитесь, что пользовательская версия Office поддерживает все API Office, которые вызывает код.

- Убедитесь в наличии определенных артефактов, таких как таблица с определенным именем.

- Назначь пользователю выбрать некоторые ячейки в Excel, а затем вставить диаграмму, инициализированную с выбранными значениями.

- Установите привязки.

- Используйте API Office диалогов, чтобы подсказыть пользователю значения параметров надстройки по умолчанию.

Однако надстройка Office успешно вызывать Office API JavaScript до загрузки библиотеки. В этой статье описываются два способа загрузки библиотеки.

- Инициализация `Office.onReady()` с .
- Инициализация `Office.initialize` с .

> [!TIP]
> Рекомендуется использовать `Office.onReady()` вместо `Office.initialize`. Хотя `Office.initialize` поддержка по-прежнему `Office.onReady()` поддерживается, обеспечивает больше гибкости. Вы можете назначить только одному обработнику, и он вызван только один раз Office `Office.initialize` инфраструктурой. Вы можете `Office.onReady()` звонить в разных местах кода и использовать различные вызовы.
> 
> Сведения о различиях описанных ниже приемов см. в статье [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Дополнительные сведения о последовательности событий при инициализации надстройки см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).

## <a name="initialize-with-officeonready"></a>Инициализация с использованием Office.onReady()

`Office.onReady()` это асинхронный метод, который возвращает объект [Promise,](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) проверяя, загружена ли Office.js библиотека. Когда библиотека загружена, она устраняет обещание как объект, который указывает Office клиентского приложения со значением enum (, и т.д.) и платформы со значением `Office.HostType` `Excel` `Word` `Office.PlatformType` enum `PC` (, `Mac` , , и `OfficeOnline` т.д.). Объект Promise сопоставляется незамедлительно, если библиотека уже загружена, когда вызывается `Office.onReady()`.

Один из способов вызова `Office.onReady()` состоит в передаче ему метода обратного вызова. Ниже приведен пример.

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

Кроме того, вы можете привязать метод `then()` к вызову `Office.onReady()`, вместо того чтобы использовать обратный вызов. Например приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Вот тот же пример с использованием ключевых `async` слов и `await` ключевых слов в TypeScript.

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, они, *как правило*, должны размещаться внутри ответа для `Office.onReady()`. Например, ссылка на [JQuery](https://jquery.com) функция `$(document).ready()` должна выглядеть следующим образом:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Однако существуют исключения для таких случаев. Например, предположим, что необходимо открыть надстройку в браузере (а не в Office приложении), чтобы отгрузить пользовательский интерфейс с помощью средств браузера. Так как Office.js не загружается в браузер, `onReady` не будет работать, а `$(document).ready` не будет работать при вызове внутри Office `onReady`. 

Другим исключением будет, если индикатор прогресса должен отображаться в области задач во время загрузки надстройки. В этом сценарии код должен вызвать jQuery и использовать его вызов для `ready` отображения индикатора прогресса. Затем обратный вызов `onReady` Office может заменять индикатор выполнения на окончательный пользовательский интерфейс  

## <a name="initialize-with-officeinitialize"></a>Инициализация с использованием Office.initialize

Событие инициализации запускается, когда библиотека Office.js будет загружена и готова к взаимодействию с пользователем. Вы можете назначить обработчик `Office.initialize` для реализации вашей логики инициализации. Например, приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Если вы используете дополнительные платформы JavaScript, которые включают собственный  обработок инициализации или тесты, они обычно должны помещаться в событие (исключения, описанные в разделе `Office.initialize` **Initialize с Office.onReady()** ранее, применяются и в этом случае). Например, ссылка на [JQuery](https://jquery.com) функция `$(document).ready()` должна выглядеть следующим образом:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Для надстроек области задач и контентных надстроек `Office.initialize` предоставляет дополнительный параметр _reason_. Этот параметр определяет порядок добавления надстройки в текущий документ. Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые или когда она уже существует в документе.

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

Дополнительные сведения см. в статьях [Событие Office.initialize](/javascript/api/office) и [Перечисление InitializationReason](/javascript/api/office/office.initializationreason).

## <a name="major-differences-between-officeinitialize-and-officeonready"></a>Основные различия между Office.initialize и Office.onReady

- Вы можете назначить только один обработчик для `Office.initialize`, который будет вызываться только один раз инфраструктурой Office, но вы можете вызывать `Office.onReady()` в разных местах вашего кода и использовать разные обратные вызовы. Например, ваш код может вызвать `Office.onReady()` сразу после загрузки настраиваемого скрипта с обратным вызовом, запускающим логику инициализации. В коде также может применяться кнопка в области задач, чей скрипт вызывает `Office.onReady()` с другим обратным вызовом. В этом случае второй обратный вызов запускается при нажатии кнопки.

- Событие `Office.initialize` запускается в конце выполнения внутренних процессов, когда Office.js инициализирует собственное выполнение. И оно срабатывает *сразу же* после окончания внутренних процессов. Если код, в котором вы назначаете обработчик события, выполняется слишком долго после запуска события, тогда обработчик не запускается. Например если вы используете диспетчер задач WebPack, он может настроить домашнюю страницу надстройки для загрузки файлов полизаполнения сразу после загрузки Office.js, но перед загрузкой вашего настраиваемого скрипта JavaScript. К тому моменту, когда ваш скрипт загружается и назначает обработчика, инициализации события уже выполнена. Но никогда не «поздно» выполнить вызов `Office.onReady()`. Если инициализация события уже произошла, обратный вызов выполняется немедленно.

> [!NOTE]
> Даже если отсутствует логика запуска, следует вызвать `Office.onReady()` или назначить пустую функцию для `Office.initialize`, когда ваша надстройка загружает JavaScript. Некоторые Office приложения и платформы не загружают области задач, пока не произойдет одно из них. Эти два способа показаны в приведенных ниже примерах.
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a>Дополнительные материалы

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md)