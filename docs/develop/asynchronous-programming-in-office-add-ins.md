---
title: Асинхронное программирование в случае надстроек Office
description: Узнайте, как Office JavaScript использует асинхронное программирование в Office надстройки.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1663f15d1b9f4191fc1f0c21f0532b5e23fdade6
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671389"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Асинхронное программирование в надстройках для Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Почему в API Надстройки Office используется асинхронное программирование? JavaScript — это язык однопотокового программирования, поэтому если скрипт вызывает продолжительный синхронный процесс, исполнение всех последующих скриптов будет заблокировано до завершения этого процесса. Поскольку некоторые операции Office веб-клиентов (но и богатых клиентов) могут блокировать выполнение, если они выполняются синхронно, большинство API JavaScript Office выполняться асинхронно. Это позволяет Office надстройки быстро и быстро. При работе с асинхронными методами зачастую требуется создавать функции обратного вызова.

Имена всех асинхронных методов в API заканчиваются "Async", такими как `Document.getSelectedDataAsync` , `Binding.getDataAsync` или `Item.loadCustomPropertiesAsync` методы. При вызове асинхронного метода он выполняется немедленно и все дополнительные скрипты могут продолжать работу. Необязательная функция обратного вызова, передаваемая в асинхронный метод, выполняется тогда, когда готовы данные или запрашиваемая операция. Обычно это происходит быстро, но иногда возможен возврат с небольшой задержкой.

На следующей схеме показан поток выполнения для вызова метода "Async", который считывает данные, выбранные пользователем в документе, открытом на сервере Word или Excel. В момент, когда выполняется вызов "Async", поток выполнения JavaScript может выполнять любую дополнительную клиентскую обработку (хотя ни один из них не отображается на схеме). Когда метод "Async" возвращается, обратное вызов возобновляет выполнение в потоке, и надстройка может получить доступ к данным, сделать что-то с ним и отобразить результат. Тот же асинхронный шаблон выполнения сохраняется при работе с Office клиентских приложений, таких как Word 2013 или Excel 2013.

*Рис. 1. Процесс выполнения при асинхронном программировании*

![Схема, показывающая взаимодействие командного выполнения со временем с пользователем, страницей надстройки и сервером веб-приложения, на котором размещена надстройка.](../images/office-addins-asynchronous-programming-flow.png)

Поддержка этой асинхронной конструкции как в полнофункциональных, так и в веб-клиентах предусмотрена в рамках стратегии проектирования "однократное написание — запуск на нескольких платформах" модели разработки надстроек Office. Например, вы можете создать надстройку области задач или контентную надстройку на единой базе кода, которая будет работать как в Excel 2013, так и в Excel в Интернете.

## <a name="writing-the-callback-function-for-an-async-method"></a>Написание функции обратного вызова для асинхронного метода

Функция вызова, которую вы  передаете в качестве аргумента вызова в метод "Async", должна объявить один параметр, который будет использовать время выполнения надстройки для предоставления доступа к объекту [AsyncResult](/javascript/api/office/office.asyncresult) при выполнении функции вызова. Можно записать:

- Анонимная функция, которая должна быть написана и передана непосредственно в  соответствии с вызовом метода "Async" в качестве параметра вызова метода "Async".

- Именоваемая функция, передав  имя этой функции в качестве параметра вызова метода "Async".

Анонимную функцию удобно использовать, если код такой функции будет использован всего один раз (так как у нее нет имени, вы не сможете сослаться на нее в другой части кода). Именованные функции применяются, если необходимо многократно использовать функцию обратного вызова для нескольких асинхронных методов.

### <a name="writing-an-anonymous-callback-function"></a>Написание анонимной функции обратного вызова

Следующая функция анонимного обратного вызова объявляет один параметр с именем, который извлекает данные из свойства `result` [AsyncResult.value](/javascript/api/office/office.asyncresult#value) при возвращении обратного вызова.

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

В следующем примере показано, как передать эту функцию анонимного вызова в строку в контексте полного вызова метода "Async" к `Document.getSelectedDataAsync` методу.

- Первый аргумент _coercionType_ указывает, чтобы вернуть выбранные данные `Office.CoercionType.Text` в виде строки текста.

- Второй аргумент _вызова_ — анонимная функция, переданная методу в строке. При выполнении функции для  отображения данных, выбранных пользователем в документе, используется параметр результатов для доступа к свойству `value` `AsyncResult` объекта.

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Вы также можете использовать параметр функции вызова для доступа к другим свойствам `AsyncResult` объекта. Используйте свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#status), чтобы определить, успешно ли был выполнен вызов. Если не удалось выполнить вызов, можно использовать свойство [AsyncResult.error](/javascript/api/office/office.asyncresult#error), чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) и получить сведения об ошибке.

Дополнительные сведения об использовании метода см. в публикации `getSelectedDataAsync` Read and write data to the active selection in a document or [spreadsheet.](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md) 

### <a name="writing-a-named-callback-function"></a>Написание именованной функции обратного вызова

Кроме того, можно написать именоваемую  функцию и передать ее имя параметру вызова метода "Async". Например, предыдущий пример можно изменить так, чтобы передавать функцию с именем `writeDataCallback` в качестве параметра _callback_.

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>Что возвращается в свойство AsyncResult.value?

Свойства объекта возвращают те же виды информации в функцию обратного вызова, переданную всем методам `asyncContext` `status` `error` `AsyncResult` async. Однако то, что возвращается в свойство, зависит от функциональности метода `AsyncResult.value` "Async".

Например, методы `addHandlerAsync` (binding, [](/javascript/api/office/office.binding) [CustomXmlPart,](/javascript/api/office/office.customxmlpart) [Document,](/javascript/api/office/office.document) [RoamingSettings](/javascript/api/outlook/office.roamingsettings)и [Параметры](/javascript/api/office/office.settings) объектов) используются для добавления функций обработки событий в элементы, представленные этими объектами. Вы можете получить доступ к свойству из функции обратного вызова, передаемой любому из методов, но так как данные или объект не доступны при добавлении обработчицы событий, свойство всегда возвращается неопределенным, если вы попытаетесь получить к нему `AsyncResult.value` `addHandlerAsync` `value` доступ. 

С другой стороны, при вызове метода он возвращает данные, выбранные пользователем в документе, в свойство `Document.getSelectedDataAsync` `AsyncResult.value` обратного вызова. Или, если вы называете метод [Bindings.getAllAsync,](/javascript/api/office/office.bindings#getAllAsync_options__callback_) он возвращает массив всех объектов `Binding` в документе. И если вы назовете метод [Bindings.getByIdAsync,](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) он возвращает один `Binding` объект.

Описание того, что возвращается в свойство для метода, см. в разделе "Значение обратного вызова" справочной `AsyncResult.value` `Async` темы этого метода. Сводку всех объектов, которые предоставляют методы, см. в таблице в нижней части темы `Async` [объекта AsyncResult.](/javascript/api/office/office.asyncresult)

## <a name="asynchronous-programming-patterns"></a>Шаблоны асинхронного программирования

API Office JavaScript поддерживает два вида асинхронных шаблонов программирования:

- С использованием вложенных обратных вызовов
- С использованием шаблона promise

При асинхронном программировании с использованием функций обратного вызова зачастую требуется вкладывать возвращаемый результат одного обратного вызова в один или несколько других обратных вызовов. В этом случае вы можете использовать вложенные обратные вызовы асинхронных методов API.

Использование вложенных обратных вызовов — это шаблон программирования, знакомый большинству разработчиков на языке JavaScript, но код с глубоко вложенными обратными вызовами может быть труден для чтения и понимания. В качестве альтернативы вложенным вызовам Office API JavaScript также поддерживает реализацию шаблона обещаний.

> [!NOTE]
> В текущей версии API Office JavaScript  встроенная поддержка шаблона обещаний работает только с кодом для привязки в Excel таблицах и документах [Word](bind-to-regions-in-a-document-or-spreadsheet.md). Однако вы можете обернуть другие функции, которые имеют обратное вызовы в вашей собственной настраиваемой функции возврата обещаний. Дополнительные сведения см. в ссылке [Wrap Common API in Promise-returning functions.](#wrap-common-apis-in-promise-returning-functions)

### <a name="asynchronous-programming-using-nested-callback-functions"></a>Асинхронное программирование с использованием вложенных функций обратного вызова

Зачастую для какой-либо задачи необходимо выполнять несколько асинхронных операций. Для этого можно вкладывать один асинхронный вызов в другой.

В следующем примере кода показано, как вложить два асинхронных вызова.

- Сначала вызывается метод [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) для получения доступа к привязке в документе с именем "MyBinding". Объект, `AsyncResult` возвращающийся к параметру обратного вызова, предоставляет доступ к указанному объекту привязки `result` из `AsyncResult.value` свойства.
- Затем объект привязки, доступный из первого параметра, используется для вызова метода `result` [Binding.getDataAsync.](/javascript/api/office/office.binding#getDataAsync_options__callback_)
- Наконец, параметр вызова, переданного методу, используется для отображения `result2` `Binding.getDataAsync` данных в привязке.

```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Этот базовый шаблон вложенного вызова можно использовать для всех асинхронных методов в Office API JavaScript.

В следующих разделах показано, как использовать анонимные или именованные функции для вложенных обратных вызовов в асинхронных методах.

#### <a name="using-anonymous-functions-for-nested-callbacks"></a>Использование анонимных функций для вложенных обратных вызовов

В следующем примере две анонимные функции объявляются inline и передаются в вложенные обратное вызовы и `getByIdAsync` `getDataAsync` методы. Поскольку это простые и встроенные функции, их назначение сразу же становится понятным.

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

#### <a name="using-named-functions-for-nested-callbacks"></a>Использование именованных функций для вложенных обратных вызовов

В сложных реализациях может оказаться полезным использовать именованные функции для упрощения чтения, поддержки и повторного использования. В следующем примере две анонимные функции из примера в предыдущем разделе были переписаны как функции с `deleteAllData` именем и `showResult` . Эти названные функции затем передаются в методы обратного вызова по `getByIdAsync` `deleteAllDataValuesAsync` имени.

```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>Асинхронное программирование с применением шаблона, предусматривающего использование обещаний для получения доступа к данным в привязках

Если применяется шаблон программирования, предусматривающий использование обещаний, в коде не нужно указывать передачу функции обратного вызова и ожидание ее возвращения для продолжения выполнения. В этом случае сразу возвращается объект обещания, который представляет нужный результат. Но в отличие от традиционного синхронного программирования, в этом случае получение обещанного результата на самом деле откладывается до тех пор, пока среда выполнения надстроек Office не сможет выполнить запрос. Обработчик _onError_ предоставляется для ситуаций, когда запрос не может быть выполнен.

API Office JavaScript предоставляет [метод Office.select](/javascript/api/office#Office_select_expression__callback_) для поддержки шаблона обещаний для работы с существующими объектами привязки. Объект promise, возвращенный методу, поддерживает только четыре метода, к которые можно получить доступ непосредственно из объекта `Office.select` [Binding:](/javascript/api/office/office.binding) [getDataAsync,](/javascript/api/office/office.binding#getDataAsync_options__callback_) [setDataAsync,](/javascript/api/office/office.binding#setDataAsync_data__options__callback_) [addHandlerAsync](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_)и [removeHandlerAsync](/javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_).

Шаблон promise для работы с привязками принимает такую форму:

**Office.select**_(selectorExpression_, _onError_**).** _BindingObjectAsyncMethod_

Параметр _selectorExpression_ принимает форму, в которой bindingId — это имя () привязки, созданной ранее в документе или таблице (с помощью одного из методов `"bindings#bindingId"`  `id` addFrom `Bindings` коллекции: `addFromNamedItemAsync` , или `addFromPromptAsync` `addFromSelectionAsync` ). Например, выражение селектора указывает, что необходимо получить доступ к привязке с `bindings#cities` **id** "cities".

Параметр _onError_ — это функция обработки ошибок, которая принимает один параметр типа, который может использоваться для доступа к объекту, если метод не может получить доступ к указанной `AsyncResult` `Error` `select` привязке. В следующем примере показана базовая функция обработки ошибки, которую можно передать в параметр _onError_.

```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Замените местообладатель _BindingObjectAsyncMethod_ вызовом на любой из четырех методов объекта, поддерживаемых объектом `Binding` promise: `getDataAsync` , , , `setDataAsync` или `addHandlerAsync` `removeHandlerAsync` . Вызовы этих методов не поддерживают дополнительные шаблоны promise. Их нужно вызывать с помощью [шаблона функции вложенного обратного вызова](#asynchronous-programming-using-nested-callback-functions).

После выполнения обещания объекта его можно повторно использовать в цепном вызове метода, как при привязке (время выполнения надстройки не будет асинхронно выполнять `Binding` обещание). Если обещание объекта не может быть выполнено, время выполнения надстройки снова будет пытаться получить доступ к объекту привязки при следующем вызове одного из его асинхронных `Binding` методов.

В следующем примере кода используется метод для получения привязки с "из коллекции", а затем вызывает метод `select` `id` `cities` `Bindings` [addHandlerAsync,](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_) чтобы добавить обработник событий для события [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) привязки.

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> Обещание `Binding` объекта, возвращаемого методом, предоставляет доступ только `Office.select` к четырем методам `Binding` объекта. Если вам нужно получить доступ к любому из других участников объекта, вместо этого необходимо использовать свойство и методы для `Binding` `Document.bindings` получения `Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding` объекта. Например, если вам необходимо получить доступ к любым свойствам объекта (свойствам или свойствам) или получить доступ к свойствам объектов `Binding` `document` `id` `type` [MatrixBinding или TableBinding,](/javascript/api/office/office.matrixbinding) [](/javascript/api/office/office.tablebinding) `getByIdAsync` `getAllAsync` необходимо использовать или методы для получения `Binding` объекта.

## <a name="passing-optional-parameters-to-asynchronous-methods"></a>Передача дополнительных параметров в асинхронные методы

Общий синтаксис методов "Async" следует следующему шаблону:

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Все асинхронные методы поддерживают дополнительные параметры, которые передаются в виде объекта JSON, содержащего один или несколько дополнительных параметров. Объект JSON, содержащий дополнительные параметры, является неупорядоченной коллекцией пар "ключ-значение" с разделителем ":". Каждая пара в объекте разделяется точкой с запятой, а весь набор пар заключен в скобки. Ключом является имя параметра, а значением — значение, которое следует передать этому параметру.

Вы можете создать объект JSON, содержащий дополнительные параметры в линию, или путем создания объекта и передачи его в качестве `options` _параметра параметра параметра параметра._

### <a name="passing-optional-parameters-inline"></a>Передача дополнительных параметров в качестве встроенных

Например, синтаксис вызова метода [Document.setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) с необязательными параметрами в качестве встроенных выглядит так:

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

В этой форме синтаксиса вызовов два необязательных параметра, _coercionType_ и _asyncContext,_ определяются как объект JSON, закрытый в скобки.

В следующем примере показано, как вызвать метод, указав дополнительные `Document.setSelectedDataAsync` параметры в линию.

```js
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> Дополнительные параметры можно задавать в объекте JSON в любом порядке, если их имена указываются правильно.

### <a name="passing-optional-parameters-in-an-options-object"></a>Передача дополнительных параметров в объекте options

Кроме того, можно создать объект с именем, который указывает необязательные параметры отдельно от вызова метода, а затем передать объект в `options` `options` качестве _аргумента параметра._

В следующем примере показан один из способов создания объекта, в котором ( и т. д.) являются задатчиками фактических имен `options` `parameter1` и `value1` значений параметров.

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

Когда указываются параметры [ValueFormat](/javascript/api/office/office.valueformat) и [FilterType](/javascript/api/office/office.filtertype), код будет таким:

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

Вот еще один способ создания `options` объекта.

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Это выглядит как следующий пример, когда используется для указания параметров и `ValueFormat` `FilterType` параметров:

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> При использовании любого метода создания объекта можно указать необязательные параметры в любом порядке, если их имена указаны `options` правильно.

В следующем примере показано, как вызвать метод, указав необязательные `Document.setSelectedDataAsync` параметры `options` объекта.

```js
var options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В обоих необязательных примерах параметров параметр _callback_ указывается как последний параметр (следуя за необязательными параметрами, или следуя объекту _аргумента параметра)._ Кроме того, параметр _callback_ можно указать либо во встроенном объекте JSON, либо в объекте `options`. Однако параметр _callback_ можно передать только одним из способов: или в объекте _options_ (встроенном или созданном внешне), или в качестве последнего параметра.

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Оберните общие API в функциях возврата обещаний

Общие методы API (Outlook API) не возвращают [обещания.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) Поэтому нельзя использовать [](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) ожидание для приостановки выполнения до завершения асинхронной операции. Если вам нужно `await` поведение, можно завернуть вызов метода в явно созданное обещание. 

Основной шаблон заключается в создании асинхронного метода, который  возвращает объект Promise немедленно и устраняет  объект Promise, когда внутренний метод завершается, или отклоняет объект, если метод не удается. Ниже приведен простой пример.

```javascript
function getDocumentFilePath() {
    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                resolve(asyncResult.value.url);
            });
        }
        catch (error) {
            reject(WordMarkdownConversion.errorHandler(error));
        }
    })
}
```

Когда этот метод необходимо ожидать, его можно назвать либо ключевым словом, либо функцией, передаемой `await` `then` функции.

> [!NOTE]
> Этот метод особенно полезен при вызове одного из общих API внутри вызова метода в одной из моделей объектов, определенных `run` приложениям. Пример вышеуказанной функции см. вHome.js в [примере Word-Add-in-JavaScript-MDConversion.](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)

Ниже приводится пример с помощью TypeScript.

```typescript
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
