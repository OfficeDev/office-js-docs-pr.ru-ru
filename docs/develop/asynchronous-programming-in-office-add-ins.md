---
title: Асинхронное программирование в случае надстроек Office
description: Узнайте, как Office JavaScript использует асинхронное программирование в Office надстроек.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c57dc1c35d518f86e4757fb1c5d6d51c9819441
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090952"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Асинхронное программирование в надстройках для Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Почему в API Надстройки Office используется асинхронное программирование? JavaScript — это язык однопотокового программирования, поэтому если скрипт вызывает продолжительный синхронный процесс, исполнение всех последующих скриптов будет заблокировано до завершения этого процесса. Так как некоторые операции с Office веб-клиентами (но также полнофункциональные клиенты) могут блокировать выполнение, если они выполняются синхронно, большинство API JavaScript Office предназначены для выполнения асинхронно. Это гарантирует, Office надстройки будут гибкими и быстрыми. При работе с асинхронными методами зачастую требуется создавать функции обратного вызова.

Имена всех асинхронных методов в API в конце "Async", `Document.getSelectedDataAsync`такие как , `Binding.getDataAsync`или методы `Item.loadCustomPropertiesAsync` . При вызове асинхронного метода он выполняется немедленно и все дополнительные скрипты могут продолжать работу. Необязательная функция обратного вызова, передаваемая в асинхронный метод, выполняется тогда, когда готовы данные или запрашиваемая операция. Обычно это происходит быстро, но иногда возможен возврат с небольшой задержкой.

На следующей схеме показан поток выполнения для вызова асинхронного метода, который считывает данные, выбранные пользователем в документе, открытом в серверном приложении Word или Excel. На момент выполнения асинхронного вызова поток выполнения JavaScript может выполнять дополнительную обработку на стороне клиента (хотя ни один из них не показан на схеме). Когда метод Async возвращается, обратный вызов возобновляет выполнение в потоке, и надстройка может получить доступ к данным, выполнить с ним что-то и отобразить результат. Тот же шаблон асинхронного выполнения используется при работе с Office клиентскими приложениями, такими как Word 2013 или Excel 2013.

*Рис. 1. Процесс выполнения при асинхронном программировании*

![Схема, показывающая взаимодействие с пользователем, страницей надстройки и сервером веб-приложений, на котором размещена надстройка.](../images/office-addins-asynchronous-programming-flow.png)

Поддержка этой асинхронной конструкции как в полнофункциональных, так и в веб-клиентах предусмотрена в рамках стратегии проектирования "однократное написание — запуск на нескольких платформах" модели разработки надстроек Office. Например, вы можете создать надстройку области задач или контентную надстройку на единой базе кода, которая будет работать как в Excel 2013, так и в Excel в Интернете.

## <a name="write-the-callback-function-for-an-async-method"></a>Написание функции обратного вызова для метода Async

Функция обратного вызова, передаваемая в качестве аргумента обратного вызова в метод Async, должна объявить один параметр, который среда выполнения надстройки будет использовать для предоставления доступа к объекту [AsyncResult](/javascript/api/office/office.asyncresult) при выполнении функции обратного вызова. Можно записать:

- Анонимная функция, которая должна быть записана и передана непосредственно в строке вызова метода Async в качестве параметра обратного вызова метода Async.

- Именованная функция, передав имя этой функции  в качестве параметра обратного вызова метода Async.

Анонимную функцию удобно использовать, если код такой функции будет использован всего один раз (так как у нее нет имени, вы не сможете сослаться на нее в другой части кода). Именованные функции применяются, если необходимо многократно использовать функцию обратного вызова для нескольких асинхронных методов.

### <a name="write-an-anonymous-callback-function"></a>Написание анонимной функции обратного вызова

Приведенная ниже анонимная `result` функция обратного вызова объявляет один параметр с именем, который извлекает данные из свойства [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) при возврате обратного вызова.

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

В следующем примере показано, как передать эту анонимную функцию обратного вызова в строке в контексте полного вызова метода Async в метод `Document.getSelectedDataAsync` .

- Первый аргумент _coercionType_ указывает, `Office.CoercionType.Text`что выбранные данные возвращаются в виде строки текста.

- Второй аргумент _обратного_ вызова — это анонимная функция, передаваемая методу в строке. При выполнении функции она использует параметр _результата_ `value` `AsyncResult` для доступа к свойству объекта для отображения данных, выбранных пользователем в документе.

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

Вы также можете использовать параметр функции обратного вызова для доступа к другим свойствам `AsyncResult` объекта. Используйте свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member), чтобы определить, успешно ли был выполнен вызов. Если не удалось выполнить вызов, можно использовать свойство [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member), чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) и получить сведения об ошибке.

Дополнительные сведения об использовании метода `getSelectedDataAsync` см. в статье "Чтение и запись данных в активное выделение в документе [или электронной таблице"](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 

### <a name="write-a-named-callback-function"></a>Написание именованной функции обратного вызова

Кроме того, можно написать именованную функцию и передать ее имя в параметр _обратного_ вызова метода Async. Например, предыдущий пример можно изменить так, чтобы передавать функцию с именем `writeDataCallback` в качестве параметра _callback_.

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

И `asyncContext`свойства `status`объекта `error` `AsyncResult` возвращают те же типы информации в функцию обратного вызова, переданную во все асинхронные методы. Однако то, что возвращается `AsyncResult.value` свойству, зависит от функциональности метода Async.

Например, `addHandlerAsync` методы (объектов [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) и [Параметры](/javascript/api/office/office.settings)) используются для добавления функций обработчика событий в элементы, представленные этими объектами. Доступ к свойству можно получить из функции обратного вызова, передаваемой в любой из методов, но так как при добавлении обработчика событий доступ к данным или объектам не выполняется, `value` свойство всегда возвращает неопределенное значение при попытке доступа к нему.`AsyncResult.value` `addHandlerAsync`

С другой стороны, при `Document.getSelectedDataAsync` вызове метода он возвращает данные, `AsyncResult.value` выбранные пользователем в документе, свойству в обратном вызове. Или при вызове метода [Bindings.getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) `Binding` он возвращает массив всех объектов в документе. При вызове метода [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) он возвращает один `Binding` объект.

Описание того, что возвращается `AsyncResult.value` `Async` свойству для метода, см. в разделе "Значение обратного вызова" справочного раздела этого метода. Сводку всех объектов `Async` , предоставляющих методы, см. в таблице в нижней части раздела объекта [AsyncResult](/javascript/api/office/office.asyncresult) .

## <a name="asynchronous-programming-patterns"></a>Шаблоны асинхронного программирования

API JavaScript Office поддерживает два типа шаблонов асинхронного программирования.

- С использованием вложенных обратных вызовов
- С использованием шаблона promise

При асинхронном программировании с использованием функций обратного вызова зачастую требуется вкладывать возвращаемый результат одного обратного вызова в один или несколько других обратных вызовов. В этом случае вы можете использовать вложенные обратные вызовы асинхронных методов API.

Использование вложенных обратных вызовов — это шаблон программирования, знакомый большинству разработчиков на языке JavaScript, но код с глубоко вложенными обратными вызовами может быть труден для чтения и понимания. В качестве альтернативы вложенным обратным вызовам Office API JavaScript также поддерживает реализацию шаблона обещаний.

> [!NOTE]
> В текущей версии API JavaScript Office встроенная поддержка шаблона promises  работает только с кодом для привязок в Excel электронных таблицах и документах [Word](bind-to-regions-in-a-document-or-spreadsheet.md). Однако можно заключить в оболочку другие функции, которые имеют обратные вызовы внутри собственной пользовательской функции, возвращаемой обещанием. Дополнительные сведения см. в статье ["Упаковка общих API-интерфейсов в функциях, возвращаемых обещанием"](#wrap-common-apis-in-promise-returning-functions).

### <a name="asynchronous-programming-using-nested-callback-functions"></a>Асинхронное программирование с использованием вложенных функций обратного вызова

Зачастую для какой-либо задачи необходимо выполнять несколько асинхронных операций. Для этого можно вкладывать один асинхронный вызов в другой.

В следующем примере кода показано, как вложить два асинхронных вызова.

- Сначала вызывается метод [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) для получения доступа к привязке в документе с именем "MyBinding". Объект `AsyncResult` , возвращаемый параметру `result` этого обратного вызова, предоставляет доступ к указанному объекту привязки из `AsyncResult.value` свойства.
- Затем объект привязки, к который был доступ из первого `result` параметра, используется для вызова метода [Binding.getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) .
- Наконец, параметр `result2` обратного вызова `Binding.getDataAsync` , переданный методу, используется для отображения данных в привязке.

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

Этот базовый шаблон вложенного обратного вызова можно использовать для всех асинхронных методов в Office API JavaScript.

В следующих разделах показано, как использовать анонимные или именованные функции для вложенных обратных вызовов в асинхронных методах.

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>Использование анонимных функций для вложенных обратных вызовов

В следующем примере две анонимные функции `getByIdAsync` объявляются встроенными и передаются в методы и в `getDataAsync` виде вложенных обратных вызовов. Поскольку это простые и встроенные функции, их назначение сразу же становится понятным.

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

#### <a name="use-named-functions-for-nested-callbacks"></a>Использование именованных функций для вложенных обратных вызовов

В сложных реализациях может оказаться полезным использовать именованные функции для упрощения чтения, поддержки и повторного использования. В следующем примере две анонимные функции из примера в предыдущем разделе были перезаписаны как функции с именем и `deleteAllData` `showResult`. Затем эти именованные функции передаются в методы `getByIdAsync` и в `deleteAllDataValuesAsync` качестве обратных вызовов по имени.

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

API Office JavaScript предоставляет [метод Office.select](/javascript/api/office#Office_select_expression__callback_) для поддержки шаблона обещаний для работы с существующими объектами привязки. Объект promise `Office.select` , возвращаемый методу, поддерживает только четыре метода, к которые можно получить доступ непосредственно из объекта [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)), [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)), [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) и [removeHandlerAsync](/javascript/api/office/office.binding#office-office-binding-removehandlerasync-member(1)).

Шаблон обещаний для работы с привязками принимает эту форму.

**Office.select(**_selectorExpression_, _onError_**).** _BindingObjectAsyncMethod_

Параметр _selectorExpression_ `"bindings#bindingId"`принимает форму, где _bindingId_ — это имя ( `id`) привязки, созданной ранее в документе или электронной таблице (с помощью одного из методов addFrom `Bindings` коллекции: `addFromNamedItemAsync`, `addFromPromptAsync`или `addFromSelectionAsync`). Например, выражение выбора указывает`bindings#cities`, что необходимо получить доступ к привязке с идентификатором "cities".

Параметр _onError_ `AsyncResult` `Error` — это функция обработки ошибок, которая принимает один параметр типа, который может использоваться для доступа к объекту, `select` если методу не удается получить доступ к указанной привязке. В следующем примере показана базовая функция обработки ошибки, которую можно передать в параметр _onError_.

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

Замените _заполнитель BindingObjectAsyncMethod_ вызовом любого из четырех `Binding` методов объекта, поддерживаемых объектом promise: `getDataAsync`, `setDataAsync`или `addHandlerAsync``removeHandlerAsync`. Вызовы этих методов не поддерживают дополнительные шаблоны promise. Их нужно вызывать с помощью [шаблона функции вложенного обратного вызова](#asynchronous-programming-using-nested-callback-functions).

После выполнения `Binding` обещания объекта его можно повторно использовать в вызове метода цепочки, как если бы это была привязка (среда выполнения надстройки не будет асинхронно повторять попытку выполнения обещания). Если не `Binding` удается выполнить обещание объекта, среда выполнения надстройки повторит попытку доступа к объекту привязки при следующем вызове одного из ее асинхронных методов.

В следующем `select` `id``cities`примере кода используется метод для получения привязки с "" `Bindings` из коллекции, а затем вызывается метод [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) для добавления обработчика событий [для события dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) привязки.

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> Обещание `Binding` объекта, возвращаемое методом `Office.select` , предоставляет доступ только к четырем методам `Binding` объекта. Если необходимо получить доступ к любому `Binding` из других элементов объекта, `Document.bindings` `Bindings.getByIdAsync` необходимо использовать свойство и методы `Bindings.getAllAsync` для извлечения `Binding` объекта. Например, `Binding` если необходимо получить доступ к любому из свойств объекта (`type``id``document`или свойств) или получить доступ к свойствам объектов [MatrixBinding](/javascript/api/office/office.matrixbinding) или [TableBinding](/javascript/api/office/office.tablebinding), `getByIdAsync` `getAllAsync` `Binding` необходимо использовать или методы для получения объекта.

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>Передача необязательных параметров в асинхронные методы

Общий синтаксис для всех методов Async следует этому шаблону.

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Все асинхронные методы поддерживают необязательные параметры, которые передаются в виде объекта JavaScript, содержащего один или несколько необязательных параметров. Объект, содержащий необязательные параметры, представляет собой неупорядоченную коллекцию пар "ключ-значение" с символом ":", разделяющим ключ и значение. Каждая пара в объекте разделяется точкой с запятой, а весь набор пар заключен в скобки. Ключом является имя параметра, а значением — значение, которое следует передать этому параметру.

Можно создать объект, содержащий встроенные необязательные параметры, `options` или создать объект и передать его в качестве _параметра_ параметров.

### <a name="pass-optional-parameters-inline"></a>Передача необязательных параметров во встроенном режиме

Например, синтаксис вызова метода [Document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) с необязательными параметрами в качестве встроенных выглядит так:

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

В этой форме вызывающего синтаксиса два необязательных параметра, _coercionType_ и _asyncContext_, определяются как анонимный объект JavaScript, встроенный в фигурные скобки.

В следующем примере показано, как вызвать метод `Document.setSelectedDataAsync` , указав необязательные встроенные параметры.

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
> Вы можете указать необязательные параметры в любом порядке в объекте параметра, если их имена указаны правильно.

### <a name="pass-optional-parameters-in-an-options-object"></a>Передача необязательных параметров в объекте options

Кроме того, можно `options` создать объект с именем, который указывает необязательные параметры отдельно от вызова метода, `options` а затем передать объект в качестве _аргумента_ параметров.

В следующем примере показан один из `options` способов создания объекта, `parameter1`где , `value1`и т. д., являются заполнителями для фактических имен и значений параметров.

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

Ниже приведен другой способ создания `options` объекта.

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Как показано в следующем примере, при использовании для указания `ValueFormat` параметров `FilterType` и параметров:

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> При использовании любого из методов `options` создания объекта можно указать необязательные параметры в любом порядке, если их имена указаны правильно.

В следующем примере показано, как вызвать метод `Document.setSelectedDataAsync` , указав необязательные параметры в объекте `options` .

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

В обоих необязательных примерах параметров параметр обратного вызова указывается в качестве последнего параметра (после встроенных необязательных параметров или после объекта _аргумента options_). Кроме того, можно указать параметр _обратного_ вызова внутри встроенного объекта JavaScript или в объекте `options` . Однако параметр _callback_ можно передать только одним из способов: или в объекте _options_ (встроенном или созданном внешне), или в качестве последнего параметра.

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Упаковка общих API-интерфейсов в функции, возвращающие обещание

Методы Common API (и Outlook API) не возвращают [promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Поэтому нельзя использовать [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) для приостановки выполнения до завершения асинхронной операции. Если требуется поведение `await` , можно заключить вызов метода в явно созданное обещание. 

Базовый шаблон — создание асинхронного метода, который немедленно возвращает объект Promise и разрешает этот  объект Promise по завершении внутреннего метода, или отклоняет объект в  случае сбоя метода. Ниже приведен простой пример.

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

Если этот метод требуется ожидать, `await` его можно вызвать с помощью ключевого слова или функции, передаваемой функции `then` .

> [!NOTE]
> Этот метод особенно полезен, если необходимо вызвать один из общих API `run` внутри вызова метода в одной из объектных моделей конкретного приложения. Пример функции выше, используемой таким образом, см. вHome.js [ примере Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).

Ниже приведен пример использования TypeScript.

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
