---
title: Инициализация надстройки Office
description: Узнайте, как инициализировать надстройку Office.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52e75770dc4852ac3905256b6ea4230552df48ca
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797598"
---
# <a name="initialize-your-office-add-in"></a>Инициализация надстройки Office

Надстройки Office часто поддерживают логику запуска для выполнения следующих действий:

- Убедитесь, что версия Office пользователя поддерживает все API Office, вызываемую кодом.

- Убедитесь, что существует определенный артефакт, например лист с определенным именем.

- Предложите пользователю выбрать некоторые ячейки в Excel, а затем вставить диаграмму, инициализированную с выбранными значениями.

- Установите привязки.

- Используйте API диалоговых окон Office, чтобы запрашивать у пользователя значения параметров надстройки по умолчанию.

Однако надстройка Office не может успешно вызывать API JavaScript для Office, пока библиотека не будет загружена. В этой статье описаны два способа загрузки библиотеки в коде.

- Инициализировать с помощью `Office.onReady()`.
- Инициализировать с помощью `Office.initialize`.

> [!TIP]
> Рекомендуется использовать `Office.onReady()` вместо `Office.initialize`. Хотя `Office.initialize` эта возможность по-прежнему поддерживается, `Office.onReady()` она обеспечивает большую гибкость. Вы можете назначить только один обработчик `Office.initialize` , и он вызывается инфраструктурой Office только один раз. Вы можете вызывать в `Office.onReady()` разных местах кода и использовать разные обратные вызовы.
> 
> Сведения о различиях описанных ниже приемов см. в статье [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).

Дополнительные сведения о последовательности событий при инициализации надстройки см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).

## <a name="initialize-with-officeonready"></a>Инициализация с использованием Office.onReady()

`Office.onReady()` — это асинхронный метод, который возвращает объект [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) во время проверки загрузки Office.js библиотеки. При загрузке библиотеки обещание разрешается как объект, указывающий клиентское приложение Office `Office.HostType` со значением перечисления (`Excel`, `Word`и т. д.) `Office.PlatformType` и платформу со значением перечисления (`PC`, `Mac`, и `OfficeOnline`т. д.). Объект Promise сопоставляется незамедлительно, если библиотека уже загружена, когда вызывается `Office.onReady()`.

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

Ниже приведен тот же пример использования ключевых `async` слов `await` и ключевых слов в TypeScript.

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

Однако существуют исключения для таких случаев. Например, предположим, что вы хотите открыть надстройку в браузере (вместо загрузки неопубликованного приложения Office) для отладки пользовательского интерфейса с помощью средств браузера. В этом сценарии, когда Office.js определяет, что оно выполняется за пределами ведущего приложения Office, `null` он вызывает обратный вызов и разрешает обещание как для узла, так и для платформы.

Другим исключением может быть отображение индикатора хода выполнения в области задач во время загрузки надстройки. В этом сценарии код должен вызвать jQuery `ready` и использовать его обратный вызов для отображения индикатора хода выполнения. Затем обратный `Office.onReady` вызов может заменить индикатор хода выполнения окончательным пользовательским интерфейсом.

## <a name="initialize-with-officeinitialize"></a>Инициализация с использованием Office.initialize

Событие инициализации запускается, когда библиотека Office.js будет загружена и готова к взаимодействию с пользователем. Вы можете назначить обработчик `Office.initialize` для реализации вашей логики инициализации. Например, приведенный ниже код проверяет, поддерживает ли версия Excel пользователя использование API, которые может вызывать надстройка.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

Если вы используете дополнительные платформы JavaScript, которые содержат собственный обработчик инициализации или  тесты, `Office.initialize` они обычно должны размещаться в событии (исключения, описанные в разделе "Инициализация с **помощью Office.onReady()** ранее применяются и в этом случае"). Например, ссылка на [JQuery](https://jquery.com) функция `$(document).ready()` должна выглядеть следующим образом:

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
> Даже если отсутствует логика запуска, следует вызвать `Office.onReady()` или назначить пустую функцию для `Office.initialize`, когда ваша надстройка загружает JavaScript. Некоторые сочетания приложений и платформ Office не загружают область задач, пока не произойдет одно из этих действий. Эти два способа показаны в приведенных ниже примерах.
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="debug-initialization"></a>Отладочная инициализация

Сведения об отладке методов `Office.initialize` и методах `Office.onReady()` см. в разделе " [Отладка методов инициализации и onReady"](../testing/debug-initialize-onready.md).

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md)