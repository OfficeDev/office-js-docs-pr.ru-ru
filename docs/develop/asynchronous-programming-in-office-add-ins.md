---
title: Асинхронное программирование в случае надстроек Office
description: Узнайте, как библиотека JavaScript для Office использует асинхронное программирование в надстройках Office.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5700ef22e9d51ab603caa84a5d329d0b56b6beca
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608448"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Асинхронное программирование в надстройках для Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Почему в API Надстройки Office используется асинхронное программирование? JavaScript — это язык однопотокового программирования, поэтому если скрипт вызывает продолжительный синхронный процесс, исполнение всех последующих скриптов будет заблокировано до завершения этого процесса. Так как определенные операции для веб-клиентов Office (но и для полнофункциональных клиентов) могут блокировать выполнение, если они выполняются синхронно, большая часть API JavaScript для Office разработана для асинхронного выполнения. Это гарантирует, что надстройки Office будут отвечать на запросы и быстро. При работе с асинхронными методами зачастую требуется создавать функции обратного вызова.

Имена всех асинхронных методов в API заканчиваются на "Async", например `Document.getSelectedDataAsync` методы, или, `Binding.getDataAsync` или `Item.loadCustomPropertiesAsync` . При вызове асинхронного метода он выполняется немедленно и все дополнительные скрипты могут продолжать работу. Необязательная функция обратного вызова, передаваемая в асинхронный метод, выполняется тогда, когда готовы данные или запрашиваемая операция. Обычно это происходит быстро, но иногда возможен возврат с небольшой задержкой.

На приведенной ниже схеме показан поток выполнения для вызова асинхронного метода, который считывает данные, выделенные пользователем в документе, открытом в серверном приложении Word или Excel. На момент вызова асинхронного метода поток выполнения JavaScript свободен для выполнения любой дополнительной обработки на стороне клиента (хотя это и не показано на схеме). Когда асинхронный метод возвращает отклик, обратный вызов возобновляет выполнение в потоке, и надстройка может получать доступ к данным, выполнять с ними операции и выводить результат. Такой же шаблон асинхронного выполнения используется при работе с ведущими приложениями полнофункционального клиента Office, например Word 2013 или Excel 2013.

*Рис. 1. Процесс выполнения при асинхронном программировании*

![Процесс выполнения асинхронного программирования](../images/office-addins-asynchronous-programming-flow.png)

Поддержка этой асинхронной конструкции как в полнофункциональных, так и в веб-клиентах предусмотрена в рамках стратегии проектирования "однократное написание — запуск на нескольких платформах" модели разработки надстроек Office. Например, вы можете создать надстройку области задач или контентную надстройку на единой базе кода, которая будет работать как в Excel 2013, так и в Excel в Интернете.

## <a name="writing-the-callback-function-for-an-async-method"></a>Написание функции обратного вызова для асинхронного метода


Функция обратного вызова, которая передается в качестве аргумента _обратного вызова_ в методе async, должна объявлять один параметр, который среда выполнения надстройки будет использовать для предоставления доступа к объекту [asyncResult](/javascript/api/office/office.asyncresult) при выполнении функции обратного вызова. Можно записать:


- Анонимная функция, которая должна быть написана и передана непосредственно в вызове асинхронного метода в качестве параметра _callback_ асинхронного метода.

- Именованная функция, передающая имя этой функции в качестве параметра _обратного вызова_ асинхронного метода.

Анонимную функцию удобно использовать, если код такой функции будет использован всего один раз (так как у нее нет имени, вы не сможете сослаться на нее в другой части кода). Именованные функции применяются, если необходимо многократно использовать функцию обратного вызова для нескольких асинхронных методов.


### <a name="writing-an-anonymous-callback-function"></a>Написание анонимной функции обратного вызова

Следующая анонимная функция обратного вызова объявляет один параметр с именем `result` , который получает данные из свойства [asyncResult. Value](/javascript/api/office/office.asyncresult#value) при возврате обратного вызова.


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

В приведенном ниже примере показано, как передать эту анонимную функцию обратного вызова в контексте полного вызова метода Async для `Document.getSelectedDataAsync` метода.


- Первый аргумент _coercionType_ , `Office.CoercionType.Text` указывает, что необходимо возвратить выбранные данные в виде строки текста.

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

Вы также можете использовать параметр функции обратного вызова для доступа к другим свойствам `AsyncResult` объекта. Используйте свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#status), чтобы определить, успешно ли был выполнен вызов. Если не удалось выполнить вызов, можно использовать свойство [AsyncResult.error](/javascript/api/office/office.asyncresult#error), чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) и получить сведения об ошибке.

Более подробную информацию об использовании `getSelectedDataAsync` метода можно узнать в [статье чтение и запись данных в активное выделение в документе или электронной таблице](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### <a name="writing-a-named-callback-function"></a>Написание именованной функции обратного вызова

Кроме того, можно написать именованную функцию и передать ее имя в параметр _callback_ асинхронного метода. Например, предыдущий пример можно изменить так, чтобы передавать функцию с именем `writeDataCallback` в качестве параметра _callback_.


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


`asyncContext`Свойства, `status` и `error` свойства `AsyncResult` объекта возвращают те же сведения в функцию обратного вызова, которая передается всем асинхронным методам. Тем не менее, возвращаемое значение `AsyncResult.value` свойства зависит от функций асинхронного метода.

Например, `addHandlerAsync` методы (для объектов [Binding](/javascript/api/office/office.binding), [CustomXMLPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [roamingSettings](/javascript/api/outlook/office.roamingsettings)и [Settings](/javascript/api/office/office.settings) ) используются для добавления функций обработчика событий к элементам, представленным этими объектами. Вы можете получить доступ к `AsyncResult.value` свойству из функции обратного вызова, которая передается любому из `addHandlerAsync` методов, но так как при попытке доступа к данным или объектам не будет выполнен доступ при добавлении обработчика событий, `value` свойство всегда возвращает значение **undefine** при попытке доступа к нему.

С другой стороны, если вызывается `Document.getSelectedDataAsync` метод, он возвращает данные, выбранные пользователем в документе, в `AsyncResult.value` свойство в обратном вызове. Или, если вызывается метод [Bindings. getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) , он возвращает массив всех `Binding` объектов в документе. При вызове метода [Bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) он возвращает один `Binding` объект.

Описание возвращаемого `AsyncResult.value` свойства для `Async` метода приведено в разделе "значение обратного вызова" раздела справки этого метода. Сводка по всем объектам, которые предоставляют `Async` методы, приведено в таблице в нижней части статьи объекта [asyncResult](/javascript/api/office/office.asyncresult) .


## <a name="asynchronous-programming-patterns"></a>Шаблоны асинхронного программирования


API JavaScript для Office поддерживает два вида шаблонов асинхронного программирования:


- С использованием вложенных обратных вызовов
    
- С использованием шаблона promise
    
При асинхронном программировании с использованием функций обратного вызова зачастую требуется вкладывать возвращаемый результат одного обратного вызова в один или несколько других обратных вызовов. В этом случае вы можете использовать вложенные обратные вызовы асинхронных методов API.

Использование вложенных обратных вызовов — это шаблон программирования, знакомый большинству разработчиков на языке JavaScript, но код с глубоко вложенными обратными вызовами может быть труден для чтения и понимания. В качестве альтернативы вложенным обратным вызовам API JavaScript для Office также поддерживает реализацию шаблона обещания. Однако в текущей версии API JavaScript для Office шаблон обещания работает только с кодом для [привязок в электронных таблицах Excel и документах Word](bind-to-regions-in-a-document-or-spreadsheet.md).

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>Асинхронное программирование с использованием вложенных функций обратного вызова


Зачастую для какой-либо задачи необходимо выполнять несколько асинхронных операций. Для этого можно вкладывать один асинхронный вызов в другой.

В следующем примере кода показано, как вложить два асинхронных вызова.


- Сначала вызывается метод [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) для получения доступа к привязке в документе с именем "MyBinding". `AsyncResult`Объект, возвращаемый `result` параметру этого обратного вызова, предоставляет доступ к указанному объекту Binding из `AsyncResult.value` Свойства.

- Затем объект привязки, к которому получен доступ из первого `result` параметра, используется для вызова метода [Binding. getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) .

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

В следующем примере две анонимные функции объявляются в виде встроенных и передаются в `getByIdAsync` методы и в `getDataAsync` качестве вложенных обратных вызовов. Поскольку это простые и встроенные функции, их назначение сразу же становится понятным.


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

В сложных реализациях может оказаться полезным использовать именованные функции для упрощения чтения, поддержки и повторного использования. В следующем примере две анонимные функции из примера, приведенного в предыдущем разделе, были переписаны как функции с именами `deleteAllData` и `showResult` . Эти именованные функции затем передаются `getByIdAsync` в `deleteAllDataValuesAsync` методы и в качестве обратных вызовов по имени.


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

Параметр _селекторекспрессион_ принимает форму `"bindings#bindingId"` , где _биндингид_ — это имя ( `id` ) привязки, созданной ранее в документе или электронной таблице (с помощью одного из методов "аддфром" `Bindings` коллекции: `addFromNamedItemAsync` , `addFromPromptAsync` или `addFromSelectionAsync` ). Например, выражение Selector `bindings#cities` указывает, что вы хотите получить доступ к привязке с **идентификатором** "городов".

Параметр _OnError_ является функцией обработки ошибок, которая принимает один параметр типа `AsyncResult` , который можно использовать для доступа к `Error` объекту, если `select` метод не может получить доступ к заданной привязке. В следующем примере показана базовая функция обработки ошибки, которую можно передать в параметр _onError_.




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

Замените заполнитель _биндингобжектасинкмесод_ на вызов любого из четырех `Binding` методов объекта, поддерживаемых объектом обещания: `getDataAsync` , `setDataAsync` , `addHandlerAsync` или `removeHandlerAsync` . Вызовы этих методов не поддерживают дополнительные шаблоны promise. Их нужно вызывать с помощью [шаблона функции вложенного обратного вызова](#AsyncProgramming_NestedCallbacks).

После выполнения `Binding` обещаний объекта его можно повторно использовать в цепочке вызовов метода, как если бы это была привязка (надстройка не будет асинхронно пытаться выполнить обещание). Если `Binding` обещание объекта не может быть выполнено, среда выполнения надстройки снова попытается получить доступ к объекту Binding при следующем вызове одного из его асинхронных методов.

В следующем примере кода используется `select` метод для получения привязки с `id` " `cities` " из `Bindings` коллекции ", а затем вызывается метод [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика событий для события [Changed](/javascript/api/office/office.bindingdatachangedeventargs) привязки.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> `Binding`Обещание объекта, возвращаемое `Office.select` методом, предоставляет доступ только к четырем методам `Binding` объекта. Если вам нужно получить доступ к любому другому элементу `Binding` объекта, необходимо использовать `Document.bindings` свойство и `Bindings.getByIdAsync` `Bindings.getAllAsync` методы для получения `Binding` объекта. Например, если необходимо получить доступ к любому `Binding` свойству объекта ( `document` `id` `type` свойствам, или свойствам) или получить доступ к свойствам объектов [MatrixBinding](/javascript/api/office/office.matrixbinding) или [TableBinding](/javascript/api/office/office.tablebinding) , необходимо использовать `getByIdAsync` `getAllAsync` методы или для получения `Binding` объекта.


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

В приведенном ниже примере показано, как создать `options` объект, где `parameter1` `value1` и т. д., представляют собой заполнители для фактических имен и значений параметров.




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

Он выглядит следующим образом при использовании для указания `ValueFormat` `FilterType` параметров and:


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> При использовании любого метода создания `options` объекта можно указать необязательные параметры в любом порядке, если их имена указываются правильно.

В приведенном ниже примере показано, как вызвать `Document.setSelectedDataAsync` метод, указав необязательные параметры в `options` объекте.




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


В примерах необязательных параметров параметр _callback_ указывается в качестве последнего параметра (после необязательных параметров, а также после объекта аргумента _Options_ ). Кроме того, параметр _callback_ можно указать либо во встроенном объекте JSON, либо в объекте `options`. Однако параметр _callback_ можно передать только одним из способов: или в объекте _options_ (встроенном или созданном внешне), или в качестве последнего параметра.


## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
