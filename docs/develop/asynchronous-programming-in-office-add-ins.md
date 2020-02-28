---
title: Асинхронное программирование в случае надстроек Office
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: fc39bddbe050f8253769a0013be2d48b26dcb599
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324647"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Асинхронное программирование в случае надстроек Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Почему API надстроек Office использует асинхронное программирование? Так как JavaScript является однопотоковым языком, если сценарий вызывает длительный синхронный процесс, все последующие сценарии будут блокироваться до завершения этого процесса. Так как определенные операции для веб-клиентов Office (но и для полнофункциональных клиентов) могут блокировать выполнение, если они выполняются синхронно, большая часть API JavaScript для Office разработана для асинхронного выполнения. Это гарантирует, что надстройки Office будут отвечать на запросы и быстро. Кроме того, при работе с этими асинхронными методами часто требуется написать функции обратного вызова.

Имена всех асинхронных методов в API заканчиваются на "Async", например методы `Document.getSelectedDataAsync`, `Binding.getDataAsync`или, или. `Item.loadCustomPropertiesAsync` При вызове асинхронного метода он выполняется немедленно и все последующие сценарии могут продолжать работу. Необязательная функция обратного вызова, которая передается асинхронному методу, выполняется сразу же после того, как данные или запрошенная операция будут готовы. Обычно это происходит быстро, но перед возвращением может быть небольшая задержка.

На приведенной ниже схеме показан поток выполнения для вызова асинхронного метода, который считывает данные, выделенные пользователем в документе, открытом в серверном приложении Word или Excel. На момент вызова асинхронного метода поток выполнения JavaScript свободен для выполнения любой дополнительной обработки на стороне клиента (хотя это и не показано на схеме). Когда асинхронный метод возвращает отклик, обратный вызов возобновляет выполнение в потоке, и надстройка может получать доступ к данным, выполнять с ними операции и выводить результат. Такой же шаблон асинхронного выполнения используется при работе с ведущими приложениями полнофункционального клиента Office, например Word 2013 или Excel 2013.

*Рис. 1. Процесс выполнения при асинхронном программировании*

![Процесс выполнения асинхронного программирования](../images/office-addins-asynchronous-programming-flow.png)

Поддержка этой асинхронной конструкции как в полнофункциональных, так и в веб-клиентах предусмотрена в рамках стратегии проектирования "однократное написание — запуск на нескольких платформах" модели разработки надстроек Office. Например, вы можете создать надстройку области задач или контентную надстройку на единой базе кода, которая будет работать как в Excel 2013, так и в Excel в Интернете.

## <a name="writing-the-callback-function-for-an-async-method"></a>Написание функции обратного вызова для асинхронного метода


Функция обратного вызова, которая передается в качестве аргумента _обратного вызова_ в методе async, должна объявлять один параметр, который среда выполнения надстройки будет использовать для предоставления доступа к объекту [asyncResult](/javascript/api/office/office.asyncresult) при выполнении функции обратного вызова. Вы можете писать:


- Анонимная функция, которая должна быть написана и передана непосредственно в вызове асинхронного метода в качестве параметра _callback_ асинхронного метода.

- Именованная функция, передающая имя этой функции в качестве параметра _обратного вызова_ асинхронного метода.

Анонимную функцию удобно использовать, если код такой функции будет использован всего один раз (так как у нее нет имени, вы не сможете сослаться на нее в другой части кода). Именованные функции применяются, если необходимо многократно использовать функцию обратного вызова для нескольких асинхронных методов.


### <a name="writing-an-anonymous-callback-function"></a>Написание анонимной функции обратного вызова

Следующая анонимная функция обратного вызова объявляет один параметр с `result` именем, который получает данные из свойства [asyncResult. Value](/javascript/api/office/office.asyncresult#value) при возврате обратного вызова.


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

В приведенном ниже примере показано, как передать эту анонимную функцию обратного вызова в контексте полного вызова метода Async для `Document.getSelectedDataAsync` метода.


- Первый аргумент _coercionType_ , `Office.CoercionType.Text`указывает, что необходимо возвратить выбранные данные в виде строки текста.

- Второй аргумент _обратного вызова_ — это анонимная функция, переданная в метод в строке. При выполнении функции она использует параметр _result_ для доступа к `value` свойству `AsyncResult` объекта для отображения данных, выбранных пользователем в документе.


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

Вы также можете использовать параметр функции обратного вызова для доступа к другим свойствам `AsyncResult` объекта. Используйте свойство [asyncResult. status](/javascript/api/office/office.asyncresult#status) , чтобы определить, успешно ли выполнен вызов или он закончился неудачно. Если при вызове произойдет сбой, можно использовать свойство [asyncResult. Error](/javascript/api/office/office.asyncresult#error) , чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) для получения сведений об ошибке.

Более подробную информацию об использовании `getSelectedDataAsync` метода можно узнать в [статье чтение и запись данных в активное выделение в документе или электронной таблице](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### <a name="writing-a-named-callback-function"></a>Написание именованной функции обратного вызова

Кроме того, можно написать именованную функцию и передать ее имя в параметр _callback_ асинхронного метода. Например, предыдущий пример можно переписать, чтобы передать функцию с именем `writeDataCallback` _обратного вызова_ , как показано ниже.


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


Свойства `asyncContext`, `status`и `error` свойства `AsyncResult` объекта возвращают те же сведения в функцию обратного вызова, которая передается всем асинхронным методам. Тем не менее, возвращаемое значение `AsyncResult.value` свойства зависит от функций асинхронного метода.

Например `addHandlerAsync` , методы (для объектов [Binding](/javascript/api/office/office.binding), [CustomXMLPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [roamingSettings](/javascript/api/outlook/office.roamingsettings)и [Settings](/javascript/api/office/office.settings) ) используются для добавления функций обработчика событий к элементам, представленным этими объектами. Вы можете получить доступ `AsyncResult.value` к свойству из функции обратного вызова, которая передается любому из `addHandlerAsync` методов, но так как при попытке доступа к данным или объектам не будет `value` выполнен доступ при добавлении обработчика событий, свойство всегда возвращает значение **undefine** при попытке доступа к нему.

С другой стороны, если вызывается `Document.getSelectedDataAsync` метод, он возвращает данные, выбранные пользователем в документе, в `AsyncResult.value` свойство в обратном вызове. Или, если вызывается метод [Bindings. getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) , он возвращает массив всех `Binding` объектов в документе. При вызове метода [Bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) он возвращает один `Binding` объект.

Описание возвращаемого `AsyncResult.value` свойства для `Async` метода приведено в разделе "значение обратного вызова" раздела справки этого метода. Сводка по всем объектам, которые предоставляют `Async` методы, приведено в таблице в нижней части статьи объекта [asyncResult](/javascript/api/office/office.asyncresult) .


## <a name="asynchronous-programming-patterns"></a>Шаблоны асинхронного программирования


API JavaScript для Office поддерживает два вида шаблонов асинхронного программирования:


- С использованием вложенных обратных вызовов
    
- С использованием шаблона promise
    
При асинхронном программировании с использованием функций обратного вызова зачастую требуется вкладывать возвращаемый результат одного обратного вызова в один или несколько других обратных вызовов. В этом случае вы можете использовать вложенные обратные вызовы асинхронных методов API.

Использование вложенных обратных вызовов — это шаблон программирования, который знаком большинству разработчиков JavaScript, но код с глубокими вложенными обратными вызовами может быть трудно читать и понимать. В качестве альтернативы вложенным обратным вызовам API JavaScript для Office также поддерживает реализацию шаблона обещания. Однако в текущей версии API JavaScript для Office шаблон обещания работает только с кодом для [привязок в электронных таблицах Excel и документах Word](bind-to-regions-in-a-document-or-spreadsheet.md).

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>Асинхронное программирование с использованием вложенных функций обратного вызова


Зачастую для какой-либо задачи необходимо выполнять несколько асинхронных операций. Для этого можно вкладывать один асинхронный вызов в другой.

В следующем примере кода показано, как вложить два асинхронных вызова.


- Сначала вызывается метод [Bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) для доступа к привязке в документе с именем "MyBinding". `AsyncResult` Объект, возвращаемый `result` параметру этого обратного вызова, предоставляет доступ к указанному объекту Binding `AsyncResult.value` из свойства.

- Затем объект привязки, к которому получен доступ из `result` первого параметра, используется для вызова метода [Binding. getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) .

- Наконец, `result2` параметр обратного вызова, передаваемый в `Binding.getDataAsync` метод, используется для отображения данных в привязке.


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

Этот базовый вложенный шаблон обратного вызова можно использовать для всех асинхронных методов в API JavaScript для Office.

В следующих разделах показано, как использовать анонимные или именованные функции для вложенных обратных вызовов в асинхронных методах.


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>Использование анонимных функций для вложенных обратных вызовов

В следующем примере две анонимные функции объявляются в виде встроенных и передаются в методы `getByIdAsync` и `getDataAsync` в качестве вложенных обратных вызовов. Так как функции просты и встроенные, цель реализации немедленно очищается.


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

В сложных реализациях может быть полезно использовать именованные функции, чтобы упростить чтение, поддержку и повторное использование кода. В следующем примере две анонимные функции из примера, приведенного в предыдущем разделе, были переписаны как функции с именами `deleteAllData` и `showResult`. Эти именованные функции затем передаются `getByIdAsync` в `deleteAllDataValuesAsync` методы и в качестве обратных вызовов по имени.


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


API JavaScript для Office предоставляет метод [Office. Select](/javascript/api/office#office-select-expression--callback-) , который поддерживает шаблон обещания для работы с существующими объектами привязки. Объект Promise, возвращенный в `Office.select` метод, поддерживает только четыре метода, к которым можно получить доступ непосредственно из объекта [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)и [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).


Шаблон promise для работы с привязками принимает такую форму:

 **Office. Select (**_Селекторекспрессион_, _OnError_**).** _Биндингобжектасинкмесод_

Параметр _селекторекспрессион_ принимает `"bindings#bindingId"`форму, где _биндингид_ — это имя ( `id`) привязки, созданной ранее в документе или электронной таблице (с помощью одного из методов "аддфром" `Bindings` коллекции: `addFromNamedItemAsync`, `addFromPromptAsync`или `addFromSelectionAsync`). Например, выражение `bindings#cities` Selector указывает, что вы хотите получить доступ к привязке с **идентификатором** "городов".

Параметр _OnError_ является функцией обработки ошибок, которая принимает один параметр типа `AsyncResult` , который можно использовать для доступа к `Error` объекту, если `select` метод не может получить доступ к заданной привязке. В следующем примере показана базовая функция обработчика ошибок, которая может быть передана в параметр _OnError_ .




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

Замените заполнитель _биндингобжектасинкмесод_ на вызов любого из четырех `Binding` методов объекта, поддерживаемых объектом обещания: `getDataAsync`, `setDataAsync`, `addHandlerAsync`или. `removeHandlerAsync` Вызовы этих методов не поддерживают дополнительные обещания. Их необходимо вызывать с помощью [вложенного шаблона функции обратного вызова](#AsyncProgramming_NestedCallbacks).

После выполнения `Binding` обещаний объекта его можно повторно использовать в цепочке вызовов метода, как если бы это была привязка (надстройка не будет асинхронно пытаться выполнить обещание). Если обещание `Binding` объекта не может быть выполнено, среда выполнения надстройки снова попытается получить доступ к объекту Binding при следующем вызове одного из его асинхронных методов.

В следующем примере кода используется `select` метод для получения привязки с `id` "" из`cities` `Bindings` коллекции ", а затем вызывается метод [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика событий для события [Changed](/javascript/api/office/office.bindingdatachangedeventargs) привязки.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> Обещание `Binding` объекта, возвращаемое `Office.select` методом, предоставляет доступ только к четырем методам `Binding` объекта. Если `Binding` вам нужно получить доступ к любому другому элементу объекта, необходимо использовать `Document.bindings` свойство и `Bindings.getByIdAsync` `Bindings.getAllAsync` методы для получения `Binding` объекта. `Binding` Например, если необходимо получить доступ к любому свойству объекта (свойствам `document`, `id`или `type` свойствам) или получить доступ к свойствам объектов [MatrixBinding](/javascript/api/office/office.matrixbinding) или [TableBinding](/javascript/api/office/office.tablebinding) , необходимо использовать методы `getByIdAsync` или `getAllAsync` для получения `Binding` объекта.


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>Передача дополнительных параметров в асинхронные методы


Общий синтаксис методов "Async" следует следующему шаблону:

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Все асинхронные методы поддерживают дополнительные параметры, которые передаются в виде объекта JSON, содержащего один или несколько дополнительных параметров. Объект JSON, содержащий дополнительные параметры, является неупорядоченной коллекцией пар "ключ-значение" с разделителем ":". Каждая пара в объекте разделяется точкой с запятой, а весь набор пар заключен в скобки. Ключом является имя параметра, а значением — значение, которое следует передать этому параметру.

Можно создать объект JSON, содержащий дополнительные встроенные параметры, или создать `options` объект и передать его в качестве параметра _Options_ .


### <a name="passing-optional-parameters-inline"></a>Передача дополнительных параметров в качестве встроенных

Например, синтаксис вызова метода [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) с необязательными параметрами в качестве встроенных выглядит так:

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

В этой форме синтаксиса вызова два необязательных параметра, _coercionType_ и _asyncContext_, ОПРЕДЕЛЯЮТся как объект JSON внутри фигурных скобок.

В приведенном ниже примере показано, как вызвать `Document.setSelectedDataAsync` метод, указав дополнительные встроенные параметры.


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

Кроме того, можно создать объект с именем `options` , который задает необязательные параметры отдельно от вызова метода, а затем передает `options` объект в качестве аргумента _Options_ .

В приведенном ниже примере показано, как создать `options` объект, где `parameter1` `value1`и т. д., представляют собой заполнители для фактических имен и значений параметров.




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

Он выглядит следующим образом при использовании для указания параметров `ValueFormat` and: `FilterType`


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> При использовании любого метода создания `options` объекта можно указать необязательные параметры в любом порядке, если их имена указываются правильно.

В приведенном ниже примере показано, как вызвать `Document.setSelectedDataAsync` метод, указав необязательные параметры `options` в объекте.




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


В примерах необязательных параметров параметр _callback_ указывается в качестве последнего параметра (после необязательных параметров, а также после объекта аргумента _Options_ ). Кроме того, можно указать параметр _обратного вызова_ в встроенном объекте JSON или в `options` объекте. Однако вы можете передать параметр _обратного вызова_ только в одном расположении: либо в объекте _Options_ (встроенном или созданном извне), либо в качестве последнего параметра, но не в обоих параметрах.


## <a name="see-also"></a>См. также

- [Общие сведения об интерфейсе API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office)
