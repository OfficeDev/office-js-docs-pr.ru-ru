---
title: Общие сведения об API JavaScript для Office
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 58829c623c06225bcc7d15925fb02a082df039c6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640094"
---
# <a name="understanding-the-javascript-api-for-office"></a>Общие сведения об API JavaScript для Office

В этой статье можно узнать об API JavaScript для Office и о том, как его использовать. Справочные сведения см. в статье [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) . О том, как обновить файлы проекта Visual Studio до последней версии API JavaScript для Office, см. в статье [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md) .

> [!NOTE]
> Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и о ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Ссылки на библиотеку API JavaScript для Office в вашей надстройке

Библиотека [API JavaScript для Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например, Excel-15.js и Outlook-15.js. Простейший способ сослаться на API — использовать нашу сеть CDN, добавив следующий код `<script>` в тег страницы `<head>`:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Это приведет к скачиванию и кэшированию файлов API JavaScript для Office при первой загрузке надстройки, чтобы убедиться, что она использует самую актуальную реализацию Office.js и сопутствующих файлов для указанной версии.

Подробные сведения о CDN-версии файла Office.js, включая способы управления версиями и обратной совместимостью, приведены в разделе [Указание ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Инициализация надстройки

**Область применения:** все типы надстроек

Надстройки Office часто имеют логику, выполняющую при запуске таких действий, как:

- Проверка того, будет ли поддерживать пользовательская версия Office все функции API для Office, вызываемые вашим кодом.

- Проверка наличия некоторых артефактов, таких как лист с конкретным именем.

- Пользователю предлагается выбрать несколько ячеек в Excel, а затем вставить диаграмму, созданную с использованием этих выбранных значений.

- Установление привязок.

- Используйте API диалога для Office, предлагающий пользователю установить для параметров надстройки значения по умолчанию.

Но ваш стартовый код не должен вызывать API-интерфейсы Office.js до тех пор, пока библиотека не будет полностью загружена. Имеется два способа для проверки вашим кодом загрузки библиотеки. Они описаны в следующих разделах: 

- [Инициализация с помощью Office.onReady()](#initialize-with-officeonready)
- [Инициализация с использованием функции Office.initialize](#initialize-with-officeinitialize)

Для получения сведений о различиях в этих методах см. [Основные различия между Office.initialize и Office.onReady()](#major-differences-between-officeinitialize-and-officeonready). Дополнительные сведения о последовательности событий при инициализации надстройки приведены в разделе [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).

### <a name="initialize-with-officeonready"></a>Инициализация с помощью Office.onReady()

`Office.onReady()` представляет собой асинхронный метод, который возвращает объект Promise, проверяя при этом, полностью ли загрузилась библиотека Office.js. Когда  библиотека загрузилась (и только тогда), метод разрешает Promise как объект, указывающий ведущее приложение Office со значением перечисления `Office.HostType` (`Excel`, `Word`и т.д.) и платформу со значением перечисления `Office.PlatformType` (`PC`, `Mac`, `OfficeOnline`, и т.д.). Если библиотека уже загружена при вызове `Office.onReady()`, объект Promise разрешается немедленно.

Один из способов вызвать `Office.onReady()` — это передать его метод обратного вызова. Ниже приведен пример:

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

Кроме того, можно объединять метод `then()` для вызова `Office.onReady()` вместо передачи обратного вызова. Например, следующий код проверяет, поддерживает ли версия Excel пользователя все API-интерфейсы, которые может вызвать надстройка.

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

Ниже приведен тот же пример с использованием ключевых слов `async` и `await` в TypeScript:

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, *как правило*, их следует размещать внутри ответа на `Office.onReady()`. Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` будет выполнена следующим образом:

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

Тем не менее, существуют исключения для этого метода. Предположим, например, что вы хотите открыть надстройку в браузере (а не загрузить ее неопубликованной в ведущем приложении Office) для отладки вашего пользовательского интерфейса с помощью инструментов веб-обозревателя. Поскольку Office.js не будет загружаться в веб-обозревателе, `onReady` и `$(document).ready` не будут выполняться при вызове внутри Office `onReady`. Еще одно исключение: вы хотите, чтобы индикатор хода выполнения отображался на панели задач в процессе загрузки надстройки. В этом сценарии код должен вызывать jQuery `ready` и использовать его обратный вызов, чтобы отобразить индикатор выполнения. Затем обратный вызов Office `onReady` сможет заменить индикатор хода выполнения окончательным вариантом пользовательского интерфейса. 

### <a name="initialize-with-officeinitialize"></a>Инициализация с использованием функции Office.initialize

Событие инициализации вызывается, когда библиотека Office.js полностью загружена и готова к взаимодействию с пользователем. Можно назначить обработчик для `Office.initialize` , который реализует логику инициализации. Ниже приведен пример, в котором проверяется, что пользователь версии Excel поддерживает все интерфейсы API, которые может вызвать надстройка.

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

При использовании дополнительных платформ JavaScript, у которых есть собственный обработчик инициализации или тесты, они должны, *как правило*, размещаться в событии `Office.initialize`. (Однако исключения, описанные ранее в разделе **Инициализация с Office.onReady()**, также применяются в этом случае.) Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` будет выполнена следующим образом:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

Для надстроек области задач и надстроек содержимого `Office.initialize` обеспечивает дополнительный параметр _reason_. Этот параметр указывает, как надстройка была добавлена в текущий документ. Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые, или когда она уже существует в документе.

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

Дополнительные сведения см. в статьях [Событие Office.initialize Event](https://docs.microsoft.com/javascript/api/office?view=office-js) и [Перечисление InitializationReason](https://docs.microsoft.com/javascript/api/office/office.initializationreason?view=office-js).

> [!NOTE]
> В настоящее время, необходимо установить `Office.Initialize`, независимо от того, вызывается ли еще и `Office.onReady()`. Если вы не используете `Office.Initialize`, вы можете задать в нем пустую функцию, как показано в следующем примере.
> 
>```js
>Office.initialize = function () {};
>```

### <a name="major-differences-between-officeinitialize-and-officeonready"></a>Основные различия между Office.initialize и Office.onReady

- Можно назначить только один обработчик для `Office.initialize`, и он вызывается только один раз в инфраструктуре Office. Однако можно вызвать `Office.onReady()` в различных местах вашего кода и использовать различные обратные вызовы. Например, код может вызывать `Office.onReady()` сразу же после того, как ваш пользовательский сценарий загрузится с помощью обратного вызова, на котором выполняется логика инициализации. У вашего кода также может быть кнопка на области задач, сценарий которой вызывает `Office.onReady()` другим обратным вызовом. В этом случае обратный вызов выполняется при нажатии кнопки.

- Событие  `Office.initialize` запускается в конце внутреннего процесса, в котором инициализируется Office.js. Оно запускается *сразу же* после завершения внутреннего процесса. Если код, в котором вы присвоили обработчика событию, выполняется слишком долго после запуска события, тогда ваш обработчик не запускается. Например, при использовании диспетчера задач WebPack он может настроить домашнюю страницу надстройки для загрузки файлов polyfill после загрузки файла Office.js, но перед загрузкой настраиваемого JavaScript. К моменту загрузки вашего сценария и назначения им обработчика событие инициализации уже произойдет. Но никогда не «слишком поздно» вызвать `Office.onReady()`. Если событие инициализации уже произошло, обратный вызов выполнится немедленно.

> [!NOTE]
> Даже если у вас нет логики запуска, вы должны назначить пустую функцию `Office.initialize` при загрузке надстройки JavaScript, как показано в следующем примере. Некоторые комбинации ведущего приложения и платформы Office не будут загружать панель задач до тех пор, пока не произойдет событие инициализации и не будет запущена указанная функция обработчика событий.
> 
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a>Объектная модель API JavaScript для Office

После инициализации надстройки могут взаимодействовать с узлом (например, Excel, Outlook). Страница [Office JavaScript API object model](office-javascript-api-object-model.md) содержит более подробные сведения по  определенному использованию шаблонов. Также имеются подробные справочные материалы по [общим API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) и конкретным узлам.

## <a name="api-support-matrix"></a>Матрица поддержки API

В этой таблице представлены API и функции, поддерживаемые всеми типами надстроек (надстройками содержимого, области задач и Outlook), а также приложения Office, в которых они могут работать, когда вы указываете ведущие приложения Office, поддерживаемые вашей надстройкой, с помощью [схемы манифестов надстроек версии 1.1 и функций, поддерживаемых API JavaScript для Office версии 1.1](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Имя узла**|База данных|Книга|Почтовый ящик|Презентация|Документ|Проект|
||**Поддерживаемые** **ведущие приложения**|Веб-приложения Access|Excel,<br/>Excel Online|Outlook,<br/>веб-приложение Outlook,<br/>OWA (веб-приложения Outlook) для устройств|PowerPoint,<br/>PowerPoint Online|Word|Project|
|**Поддерживаемые типы надстроек**|Содержимое|Да|Да||Да|||
||Область задач||Да||Да|Да|Да|
||Outlook|||Да||||
|**Поддерживаемые функции API**|Чтение/запись текста||Да||Да|Да|Да<br/>(только для чтения)|
||Чтение/запись матрицы||Да|||Да||
||Чтение/запись таблицы||Да|||Да||
||Чтение/запись HTML|||||Да||
||Чтение/запись<br/>Office Open XML|||||Да||
||Чтение свойств task, resource, view и field||||||Да|
||События изменения выделения||Да|||Да||
||Загрузка всего документа||||Да|Да||
||Привязки и их события|Да<br/>(только полные и частичные привязки таблиц)|Да|||Да||
||Чтение/запись настраиваемых XML-частей|||||Да||
||Сохранение данных состояния надстройки (параметры)|Да<br/>(на ведущую надстройку)|Да<br/>(на документ)|Да<br/>(на почтовый ящик)|Да<br/>(на документ)|Да<br/>(на документ)||
||События изменения параметров|Да|Да||Да|Да||
||Получение активного режима просмотра<br/>и просмотр измененных событий||||Да|||
||Переход к расположениям<br/>в документе||Да||Да|Да||
||Активация в зависимости от контекста<br/>с помощью правил и регулярных выражений|||Да||||
||Чтение свойств элемента|||Да||||
||Чтение профиля пользователя|||Да||||
||Получение вложений|||Да||||
||Получение маркера удостоверения пользователя|||Да||||
||Вызов веб-служб Exchange|||Да||||
