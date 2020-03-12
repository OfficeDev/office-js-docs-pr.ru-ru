---
title: Асинхронное программирование в случае надстроек Office
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 931ef17115885c8f96d41bf00143b3269a515d56
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596692"
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="f2725-102">Асинхронное программирование в случае надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f2725-102">Asynchronous programming in Office Add-ins</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="f2725-p101">Почему API надстроек Office использует асинхронное программирование? Так как JavaScript является однопотоковым языком, если сценарий вызывает длительный синхронный процесс, все последующие сценарии будут блокироваться до завершения этого процесса. Так как определенные операции для веб-клиентов Office (но и для полнофункциональных клиентов) могут блокировать выполнение, если они выполняются синхронно, большая часть API JavaScript для Office разработана для асинхронного выполнения. Это гарантирует, что надстройки Office будут отвечать на запросы и быстро. Кроме того, при работе с этими асинхронными методами часто требуется написать функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f2725-p101">Why does the Office Add-ins API use asynchronous programming? Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes. Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the Office JavaScript APIs are designed to execute asynchronously. This makes sure that Office Add-ins are responsive and fast. It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="f2725-p102">Имена всех асинхронных методов в API заканчиваются на "Async", например методы `Document.getSelectedDataAsync`, `Binding.getDataAsync`или, или. `Item.loadCustomPropertiesAsync` При вызове асинхронного метода он выполняется немедленно и все последующие сценарии могут продолжать работу. Необязательная функция обратного вызова, которая передается асинхронному методу, выполняется сразу же после того, как данные или запрошенная операция будут готовы. Обычно это происходит быстро, но перед возвращением может быть небольшая задержка.</span><span class="sxs-lookup"><span data-stu-id="f2725-p102">The names of all asynchronous methods in the API end with "Async", such as the `Document.getSelectedDataAsync`, `Binding.getDataAsync`, or `Item.loadCustomPropertiesAsync` methods. When an "Async" method is called, it executes immediately and any subsequent script execution can continue. The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready. This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="f2725-p103">На приведенной ниже схеме показан поток выполнения для вызова асинхронного метода, который считывает данные, выделенные пользователем в документе, открытом в серверном приложении Word или Excel. На момент вызова асинхронного метода поток выполнения JavaScript свободен для выполнения любой дополнительной обработки на стороне клиента (хотя это и не показано на схеме). Когда асинхронный метод возвращает отклик, обратный вызов возобновляет выполнение в потоке, и надстройка может получать доступ к данным, выполнять с ними операции и выводить результат. Такой же шаблон асинхронного выполнения используется при работе с ведущими приложениями полнофункционального клиента Office, например Word 2013 или Excel 2013.</span><span class="sxs-lookup"><span data-stu-id="f2725-p103">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word or Excel. At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing (although none are shown in the diagram). When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result. The same asynchronous execution pattern holds when working with the Office rich client host applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="f2725-116">*Рис. 1. Процесс выполнения при асинхронном программировании*</span><span class="sxs-lookup"><span data-stu-id="f2725-116">*Figure 1. Asynchronous programming execution flow*</span></span>

![Процесс выполнения асинхронного программирования](../images/office-addins-asynchronous-programming-flow.png)

<span data-ttu-id="f2725-p104">Поддержка этой асинхронной конструкции как в полнофункциональных, так и в веб-клиентах предусмотрена в рамках стратегии проектирования "однократное написание — запуск на нескольких платформах" модели разработки надстроек Office. Например, вы можете создать надстройку области задач или контентную надстройку на единой базе кода, которая будет работать как в Excel 2013, так и в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="f2725-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel on the web.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="f2725-120">Написание функции обратного вызова для асинхронного метода</span><span class="sxs-lookup"><span data-stu-id="f2725-120">Writing the callback function for an "Async" method</span></span>


<span data-ttu-id="f2725-p105">Функция обратного вызова, которая передается в качестве аргумента _обратного вызова_ в методе async, должна объявлять один параметр, который среда выполнения надстройки будет использовать для предоставления доступа к объекту [asyncResult](/javascript/api/office/office.asyncresult) при выполнении функции обратного вызова. Вы можете писать:</span><span class="sxs-lookup"><span data-stu-id="f2725-p105">The callback function you pass as the _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](/javascript/api/office/office.asyncresult) object when the callback function executes. You can write:</span></span>


- <span data-ttu-id="f2725-123">Анонимная функция, которая должна быть написана и передана непосредственно в вызове асинхронного метода в качестве параметра _callback_ асинхронного метода.</span><span class="sxs-lookup"><span data-stu-id="f2725-123">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the _callback_ parameter of the "Async" method.</span></span>

- <span data-ttu-id="f2725-124">Именованная функция, передающая имя этой функции в качестве параметра _обратного вызова_ асинхронного метода.</span><span class="sxs-lookup"><span data-stu-id="f2725-124">A named function, passing the name of that function as the _callback_ parameter of an "Async" method.</span></span>

<span data-ttu-id="f2725-p106">Анонимную функцию удобно использовать, если код такой функции будет использован всего один раз (так как у нее нет имени, вы не сможете сослаться на нее в другой части кода). Именованные функции применяются, если необходимо многократно использовать функцию обратного вызова для нескольких асинхронных методов.</span><span class="sxs-lookup"><span data-stu-id="f2725-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>


### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="f2725-127">Написание анонимной функции обратного вызова</span><span class="sxs-lookup"><span data-stu-id="f2725-127">Writing an anonymous callback function</span></span>

<span data-ttu-id="f2725-128">Следующая анонимная функция обратного вызова объявляет один параметр с `result` именем, который получает данные из свойства [asyncResult. Value](/javascript/api/office/office.asyncresult#value) при возврате обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f2725-128">The following anonymous callback function declares a single parameter named `result` that retrieves data from the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property when the callback returns.</span></span>


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="f2725-129">В приведенном ниже примере показано, как передать эту анонимную функцию обратного вызова в контексте полного вызова метода Async для `Document.getSelectedDataAsync` метода.</span><span class="sxs-lookup"><span data-stu-id="f2725-129">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the `Document.getSelectedDataAsync` method.</span></span>


- <span data-ttu-id="f2725-130">Первый аргумент _coercionType_ , `Office.CoercionType.Text`указывает, что необходимо возвратить выбранные данные в виде строки текста.</span><span class="sxs-lookup"><span data-stu-id="f2725-130">The first _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>

- <span data-ttu-id="f2725-p107">Второй аргумент _обратного вызова_ — это анонимная функция, переданная в метод в строке. При выполнении функции она использует параметр _result_ для доступа к `value` свойству `AsyncResult` объекта для отображения данных, выбранных пользователем в документе.</span><span class="sxs-lookup"><span data-stu-id="f2725-p107">The second _callback_ argument is the anonymous function passed in-line to the method. When the function executes, it uses the _result_ parameter to access the `value` property of the `AsyncResult` object to display the data selected by the user in the document.</span></span>


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

<span data-ttu-id="f2725-p108">Вы также можете использовать параметр функции обратного вызова для доступа к другим свойствам `AsyncResult` объекта. Используйте свойство [asyncResult. status](/javascript/api/office/office.asyncresult#status) , чтобы определить, успешно ли выполнен вызов или он закончился неудачно. Если при вызове произойдет сбой, можно использовать свойство [asyncResult. Error](/javascript/api/office/office.asyncresult#error) , чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) для получения сведений об ошибке.</span><span class="sxs-lookup"><span data-stu-id="f2725-p108">You can also use the parameter of your callback function to access other properties of the `AsyncResult` object. Use the [AsyncResult.status](/javascript/api/office/office.asyncresult#status) property to determine if the call succeeded or failed. If your call fails you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#error) property to access an [Error](/javascript/api/office/office.error) object for error information.</span></span>

<span data-ttu-id="f2725-136">Более подробную информацию об использовании `getSelectedDataAsync` метода можно узнать в [статье чтение и запись данных в активное выделение в документе или электронной таблице](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="f2725-136">For more information about using the `getSelectedDataAsync` method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 


### <a name="writing-a-named-callback-function"></a><span data-ttu-id="f2725-137">Написание именованной функции обратного вызова</span><span class="sxs-lookup"><span data-stu-id="f2725-137">Writing a named callback function</span></span>

<span data-ttu-id="f2725-p109">Кроме того, можно написать именованную функцию и передать ее имя в параметр _callback_ асинхронного метода. Например, предыдущий пример можно переписать, чтобы передать функцию с именем `writeDataCallback` _обратного вызова_ , как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="f2725-p109">Alternatively, you can write a named function and pass its name to the _callback_ parameter of an "Async" method. For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>


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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="f2725-140">Что возвращается в свойство AsyncResult.value?</span><span class="sxs-lookup"><span data-stu-id="f2725-140">Differences in what's returned to the AsyncResult.value property</span></span>


<span data-ttu-id="f2725-p110">Свойства `asyncContext`, `status`и `error` свойства `AsyncResult` объекта возвращают те же сведения в функцию обратного вызова, которая передается всем асинхронным методам. Тем не менее, возвращаемое значение `AsyncResult.value` свойства зависит от функций асинхронного метода.</span><span class="sxs-lookup"><span data-stu-id="f2725-p110">The `asyncContext`, `status`, and `error` properties of the `AsyncResult` object return the same kinds of information to the callback function passed to all "Async" methods. However, what's returned to the `AsyncResult.value` property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="f2725-p111">Например `addHandlerAsync` , методы (для объектов [Binding](/javascript/api/office/office.binding), [CustomXMLPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [roamingSettings](/javascript/api/outlook/office.roamingsettings)и [Settings](/javascript/api/office/office.settings) ) используются для добавления функций обработчика событий к элементам, представленным этими объектами. Вы можете получить доступ `AsyncResult.value` к свойству из функции обратного вызова, которая передается любому из `addHandlerAsync` методов, но так как при попытке доступа к данным или объектам не будет `value` выполнен доступ при добавлении обработчика событий, свойство всегда возвращает значение **undefine** при попытке доступа к нему.</span><span class="sxs-lookup"><span data-stu-id="f2725-p111">For example, the `addHandlerAsync` methods (of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [Settings](/javascript/api/office/office.settings) objects) are used to add event handler functions to the items represented by these objects. You can access the `AsyncResult.value` property from the callback function you pass to any of the `addHandlerAsync` methods, but since no data or object is being accessed when you add an event handler, the `value` property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="f2725-p112">С другой стороны, если вызывается `Document.getSelectedDataAsync` метод, он возвращает данные, выбранные пользователем в документе, в `AsyncResult.value` свойство в обратном вызове. Или, если вызывается метод [Bindings. getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) , он возвращает массив всех `Binding` объектов в документе. При вызове метода [Bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) он возвращает один `Binding` объект.</span><span class="sxs-lookup"><span data-stu-id="f2725-p112">On the other hand, if you call the `Document.getSelectedDataAsync` method, it returns the data the user selected in the document to the `AsyncResult.value` property in the callback. Or, if you call the [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) method, it returns an array of all of the `Binding` objects in the document. And, if you call the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method, it returns a single `Binding` object.</span></span>

<span data-ttu-id="f2725-p113">Описание возвращаемого `AsyncResult.value` свойства для `Async` метода приведено в разделе "значение обратного вызова" раздела справки этого метода. Сводка по всем объектам, которые предоставляют `Async` методы, приведено в таблице в нижней части статьи объекта [asyncResult](/javascript/api/office/office.asyncresult) .</span><span class="sxs-lookup"><span data-stu-id="f2725-p113">For a description of what's returned to the `AsyncResult.value` property for an `Async` method, see the "Callback value" section of that method's reference topic. For a summary of all of the objects that provide `Async` methods, see the table at the bottom of the [AsyncResult](/javascript/api/office/office.asyncresult) object topic.</span></span>


## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="f2725-150">Шаблоны асинхронного программирования</span><span class="sxs-lookup"><span data-stu-id="f2725-150">Asynchronous programming patterns</span></span>


<span data-ttu-id="f2725-151">API JavaScript для Office поддерживает два вида шаблонов асинхронного программирования:</span><span class="sxs-lookup"><span data-stu-id="f2725-151">The Office JavaScript API supports two kinds of asynchronous programming patterns:</span></span>


- <span data-ttu-id="f2725-152">С использованием вложенных обратных вызовов</span><span class="sxs-lookup"><span data-stu-id="f2725-152">Using nested callbacks</span></span>
    
- <span data-ttu-id="f2725-153">С использованием шаблона promise</span><span class="sxs-lookup"><span data-stu-id="f2725-153">Using the promises pattern</span></span>
    
<span data-ttu-id="f2725-p114">При асинхронном программировании с использованием функций обратного вызова зачастую требуется вкладывать возвращаемый результат одного обратного вызова в один или несколько других обратных вызовов. В этом случае вы можете использовать вложенные обратные вызовы асинхронных методов API.</span><span class="sxs-lookup"><span data-stu-id="f2725-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="f2725-p115">Использование вложенных обратных вызовов — это шаблон программирования, который знаком большинству разработчиков JavaScript, но код с глубокими вложенными обратными вызовами может быть трудно читать и понимать. В качестве альтернативы вложенным обратным вызовам API JavaScript для Office также поддерживает реализацию шаблона обещания. Однако в текущей версии API JavaScript для Office шаблон обещания работает только с кодом для [привязок в электронных таблицах Excel и документах Word](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="f2725-p115">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand. As an alternative to nested callbacks, the Office JavaScript API also supports an implementation of the promises pattern. However, in the current version of the Office JavaScript API, the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="f2725-159">Асинхронное программирование с использованием вложенных функций обратного вызова</span><span class="sxs-lookup"><span data-stu-id="f2725-159">Asynchronous programming using nested callback functions</span></span>


<span data-ttu-id="f2725-p116">Зачастую для какой-либо задачи необходимо выполнять несколько асинхронных операций. Для этого можно вкладывать один асинхронный вызов в другой.</span><span class="sxs-lookup"><span data-stu-id="f2725-p116">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span>

<span data-ttu-id="f2725-162">В следующем примере кода показано, как вложить два асинхронных вызова.</span><span class="sxs-lookup"><span data-stu-id="f2725-162">The following code example nests two asynchronous calls.</span></span>


- <span data-ttu-id="f2725-p117">Сначала вызывается метод [Bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) для доступа к привязке в документе с именем "MyBinding". `AsyncResult` Объект, возвращаемый `result` параметру этого обратного вызова, предоставляет доступ к указанному объекту Binding `AsyncResult.value` из свойства.</span><span class="sxs-lookup"><span data-stu-id="f2725-p117">First, the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method is called to access a binding in the document named "MyBinding". The `AsyncResult` object returned to the `result` parameter of that callback provides access to the specified binding object from the `AsyncResult.value` property.</span></span>

- <span data-ttu-id="f2725-165">Затем объект привязки, к которому получен доступ из `result` первого параметра, используется для вызова метода [Binding. getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="f2725-165">Then, the binding object accessed from the first `result` parameter is used to call the [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) method.</span></span>

- <span data-ttu-id="f2725-166">Наконец, `result2` параметр обратного вызова, передаваемый в `Binding.getDataAsync` метод, используется для отображения данных в привязке.</span><span class="sxs-lookup"><span data-stu-id="f2725-166">Finally, the `result2` parameter of the callback passed to the `Binding.getDataAsync` method is used to display the data in the binding.</span></span>


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

<span data-ttu-id="f2725-167">Этот базовый вложенный шаблон обратного вызова можно использовать для всех асинхронных методов в API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="f2725-167">This basic nested callback pattern can be used for all asynchronous methods in the Office JavaScript API.</span></span>

<span data-ttu-id="f2725-168">В следующих разделах показано, как использовать анонимные или именованные функции для вложенных обратных вызовов в асинхронных методах.</span><span class="sxs-lookup"><span data-stu-id="f2725-168">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>


#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="f2725-169">Использование анонимных функций для вложенных обратных вызовов</span><span class="sxs-lookup"><span data-stu-id="f2725-169">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="f2725-p118">В следующем примере две анонимные функции объявляются в виде встроенных и передаются в методы `getByIdAsync` и `getDataAsync` в качестве вложенных обратных вызовов. Так как функции просты и встроенные, цель реализации немедленно очищается.</span><span class="sxs-lookup"><span data-stu-id="f2725-p118">In the following example, two anonymous functions are declared inline and passed into the `getByIdAsync` and `getDataAsync` methods as nested callbacks. Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>


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


#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="f2725-172">Использование именованных функций для вложенных обратных вызовов</span><span class="sxs-lookup"><span data-stu-id="f2725-172">Using named functions for nested callbacks</span></span>

<span data-ttu-id="f2725-p119">В сложных реализациях может быть полезно использовать именованные функции, чтобы упростить чтение, поддержку и повторное использование кода. В следующем примере две анонимные функции из примера, приведенного в предыдущем разделе, были переписаны как функции с именами `deleteAllData` и `showResult`. Эти именованные функции затем передаются `getByIdAsync` в `deleteAllDataValuesAsync` методы и в качестве обратных вызовов по имени.</span><span class="sxs-lookup"><span data-stu-id="f2725-p119">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse. In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named `deleteAllData` and `showResult`. These named functions are then passed into the `getByIdAsync` and `deleteAllDataValuesAsync` methods as callbacks by name.</span></span>


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


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="f2725-176">Асинхронное программирование с применением шаблона, предусматривающего использование обещаний для получения доступа к данным в привязках</span><span class="sxs-lookup"><span data-stu-id="f2725-176">Asynchronous programming using the promises pattern to access data in bindings</span></span>


<span data-ttu-id="f2725-p120">Если применяется шаблон программирования, предусматривающий использование обещаний, в коде не нужно указывать передачу функции обратного вызова и ожидание ее возвращения для продолжения выполнения. В этом случае сразу возвращается объект обещания, который представляет нужный результат. Но в отличие от традиционного синхронного программирования, в этом случае получение обещанного результата на самом деле откладывается до тех пор, пока среда выполнения надстроек Office не сможет выполнить запрос. Обработчик _onError_ предоставляется для ситуаций, когда запрос не может быть выполнен.</span><span class="sxs-lookup"><span data-stu-id="f2725-p120">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>


<span data-ttu-id="f2725-p121">API JavaScript для Office предоставляет метод [Office. Select](/javascript/api/office#office-select-expression--callback-) , который поддерживает шаблон обещания для работы с существующими объектами привязки. Объект Promise, возвращенный в `Office.select` метод, поддерживает только четыре метода, к которым можно получить доступ непосредственно из объекта [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)и [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="f2725-p121">The Office JavaScript API provides the [Office.select](/javascript/api/office#office-select-expression--callback-) method to support the promises pattern for working with existing binding objects. The promise object returned to the `Office.select` method supports only the four methods that you can access directly from the [Binding](/javascript/api/office/office.binding) object: [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-), and [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span></span>


<span data-ttu-id="f2725-182">Шаблон promise для работы с привязками принимает такую форму:</span><span class="sxs-lookup"><span data-stu-id="f2725-182">The promises pattern for working with bindings takes this form:</span></span>

 <span data-ttu-id="f2725-183">**Office. Select (**_Селекторекспрессион_, _OnError_**).** _Биндингобжектасинкмесод_</span><span class="sxs-lookup"><span data-stu-id="f2725-183">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="f2725-p122">Параметр _селекторекспрессион_ принимает `"bindings#bindingId"`форму, где _биндингид_ — это имя ( `id`) привязки, созданной ранее в документе или электронной таблице (с помощью одного из методов "аддфром" `Bindings` коллекции: `addFromNamedItemAsync`, `addFromPromptAsync`или `addFromSelectionAsync`). Например, выражение `bindings#cities` Selector указывает, что вы хотите получить доступ к привязке с **идентификатором** "городов".</span><span class="sxs-lookup"><span data-stu-id="f2725-p122">The _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where _bindingId_ is the name ( `id`) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the `Bindings` collection: `addFromNamedItemAsync`, `addFromPromptAsync`, or `addFromSelectionAsync`). For example, the selector expression `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="f2725-p123">Параметр _OnError_ является функцией обработки ошибок, которая принимает один параметр типа `AsyncResult` , который можно использовать для доступа к `Error` объекту, если `select` метод не может получить доступ к заданной привязке. В следующем примере показана базовая функция обработчика ошибок, которая может быть передана в параметр _OnError_ .</span><span class="sxs-lookup"><span data-stu-id="f2725-p123">The _onError_ parameter is an error handling function which takes a single parameter of type `AsyncResult` that can be used to access an `Error` object, if the `select` method fails to access the specified binding. The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>




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

<span data-ttu-id="f2725-p124">Замените заполнитель _биндингобжектасинкмесод_ на вызов любого из четырех `Binding` методов объекта, поддерживаемых объектом обещания: `getDataAsync`, `setDataAsync`, `addHandlerAsync`или. `removeHandlerAsync` Вызовы этих методов не поддерживают дополнительные обещания. Их необходимо вызывать с помощью [вложенного шаблона функции обратного вызова](#AsyncProgramming_NestedCallbacks).</span><span class="sxs-lookup"><span data-stu-id="f2725-p124">Replace the _BindingObjectAsyncMethod_ placeholder with a call to any of the four `Binding` object methods supported by the promise object: `getDataAsync`, `setDataAsync`, `addHandlerAsync`, or `removeHandlerAsync`. Calls to these methods don't support additional promises. You must call them using the [nested callback function pattern](#AsyncProgramming_NestedCallbacks).</span></span>

<span data-ttu-id="f2725-p125">После выполнения `Binding` обещаний объекта его можно повторно использовать в цепочке вызовов метода, как если бы это была привязка (надстройка не будет асинхронно пытаться выполнить обещание). Если обещание `Binding` объекта не может быть выполнено, среда выполнения надстройки снова попытается получить доступ к объекту Binding при следующем вызове одного из его асинхронных методов.</span><span class="sxs-lookup"><span data-stu-id="f2725-p125">After a `Binding` object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise). If the `Binding` object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="f2725-193">В следующем примере кода используется `select` метод для получения привязки с `id` "" из`cities` `Bindings` коллекции ", а затем вызывается метод [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) для добавления обработчика событий для события [Changed](/javascript/api/office/office.bindingdatachangedeventargs) привязки.</span><span class="sxs-lookup"><span data-stu-id="f2725-193">The following code example uses the `select` method to retrieve a binding with the `id` "`cities`" from the `Bindings` collection, and then calls the [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event of the binding.</span></span>




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> <span data-ttu-id="f2725-p126">Обещание `Binding` объекта, возвращаемое `Office.select` методом, предоставляет доступ только к четырем методам `Binding` объекта. Если `Binding` вам нужно получить доступ к любому другому элементу объекта, необходимо использовать `Document.bindings` свойство и `Bindings.getByIdAsync` `Bindings.getAllAsync` методы для получения `Binding` объекта. `Binding` Например, если необходимо получить доступ к любому свойству объекта (свойствам `document`, `id`или `type` свойствам) или получить доступ к свойствам объектов [MatrixBinding](/javascript/api/office/office.matrixbinding) или [TableBinding](/javascript/api/office/office.tablebinding) , необходимо использовать методы `getByIdAsync` или `getAllAsync` для получения `Binding` объекта.</span><span class="sxs-lookup"><span data-stu-id="f2725-p126">The `Binding` object promise returned by the `Office.select` method provides access to only the four methods of the `Binding` object. If you need to access any of the other members of the `Binding` object, instead you must use the `Document.bindings` property and `Bindings.getByIdAsync` or `Bindings.getAllAsync` methods to retrieve the `Binding` object. For example, if you need to access any of the `Binding` object's properties (the `document`, `id`, or `type` properties), or need to access the properties of the [MatrixBinding](/javascript/api/office/office.matrixbinding) or [TableBinding](/javascript/api/office/office.tablebinding) objects, you must use the `getByIdAsync` or `getAllAsync` methods to retrieve a `Binding` object.</span></span>


## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="f2725-197">Передача дополнительных параметров в асинхронные методы</span><span class="sxs-lookup"><span data-stu-id="f2725-197">Passing optional parameters to asynchronous methods</span></span>


<span data-ttu-id="f2725-198">Общий синтаксис методов "Async" следует следующему шаблону:</span><span class="sxs-lookup"><span data-stu-id="f2725-198">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="f2725-199">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`</span><span class="sxs-lookup"><span data-stu-id="f2725-199">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="f2725-p127">Все асинхронные методы поддерживают дополнительные параметры, которые передаются в виде объекта JSON, содержащего один или несколько дополнительных параметров. Объект JSON, содержащий дополнительные параметры, является неупорядоченной коллекцией пар "ключ-значение" с разделителем ":". Каждая пара в объекте разделяется точкой с запятой, а весь набор пар заключен в скобки. Ключом является имя параметра, а значением — значение, которое следует передать этому параметру.</span><span class="sxs-lookup"><span data-stu-id="f2725-p127">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="f2725-204">Можно создать объект JSON, содержащий дополнительные встроенные параметры, или создать `options` объект и передать его в качестве параметра _Options_ .</span><span class="sxs-lookup"><span data-stu-id="f2725-204">You can create the JSON object that contains optional parameters inline, or by creating an `options` object and passing that in as the _options_ parameter.</span></span>


### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="f2725-205">Передача дополнительных параметров в качестве встроенных</span><span class="sxs-lookup"><span data-stu-id="f2725-205">Passing optional parameters inline</span></span>

<span data-ttu-id="f2725-206">Например, синтаксис вызова метода [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) с необязательными параметрами в качестве встроенных выглядит так:</span><span class="sxs-lookup"><span data-stu-id="f2725-206">For example, the syntax for calling the [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

<span data-ttu-id="f2725-207">В этой форме синтаксиса вызова два необязательных параметра, _coercionType_ и _asyncContext_, ОПРЕДЕЛЯЮТся как объект JSON внутри фигурных скобок.</span><span class="sxs-lookup"><span data-stu-id="f2725-207">In this form of the calling syntax, the two optional parameters, _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="f2725-208">В приведенном ниже примере показано, как вызвать `Document.setSelectedDataAsync` метод, указав дополнительные встроенные параметры.</span><span class="sxs-lookup"><span data-stu-id="f2725-208">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters inline.</span></span>


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
> <span data-ttu-id="f2725-209">Дополнительные параметры можно задавать в объекте JSON в любом порядке, если их имена указываются правильно.</span><span class="sxs-lookup"><span data-stu-id="f2725-209">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>


### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="f2725-210">Передача дополнительных параметров в объекте options</span><span class="sxs-lookup"><span data-stu-id="f2725-210">Passing optional parameters in an options object</span></span>

<span data-ttu-id="f2725-211">Кроме того, можно создать объект с именем `options` , который задает необязательные параметры отдельно от вызова метода, а затем передает `options` объект в качестве аргумента _Options_ .</span><span class="sxs-lookup"><span data-stu-id="f2725-211">Alternatively, you can create an object named `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="f2725-212">В приведенном ниже примере показано, как создать `options` объект, где `parameter1` `value1`и т. д., представляют собой заполнители для фактических имен и значений параметров.</span><span class="sxs-lookup"><span data-stu-id="f2725-212">The following example shows one way of creating the `options` object, where `parameter1`, `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="f2725-213">Когда указываются параметры [ValueFormat](/javascript/api/office/office.valueformat) и [FilterType](/javascript/api/office/office.filtertype), код будет таким:</span><span class="sxs-lookup"><span data-stu-id="f2725-213">Which looks like the following example when used to specify the [ValueFormat](/javascript/api/office/office.valueformat) and [FilterType](/javascript/api/office/office.filtertype) parameters.</span></span>




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="f2725-214">Вот еще один способ создания `options` объекта.</span><span class="sxs-lookup"><span data-stu-id="f2725-214">Here's another way of creating the `options` object.</span></span>




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="f2725-215">Он выглядит следующим образом при использовании для указания параметров `ValueFormat` and: `FilterType`</span><span class="sxs-lookup"><span data-stu-id="f2725-215">Which looks like the following example when used to specify the `ValueFormat` and `FilterType` parameters:</span></span>


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> <span data-ttu-id="f2725-216">При использовании любого метода создания `options` объекта можно указать необязательные параметры в любом порядке, если их имена указываются правильно.</span><span class="sxs-lookup"><span data-stu-id="f2725-216">When using either method of creating the `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="f2725-217">В приведенном ниже примере показано, как вызвать `Document.setSelectedDataAsync` метод, указав необязательные параметры `options` в объекте.</span><span class="sxs-lookup"><span data-stu-id="f2725-217">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters in an `options` object.</span></span>




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


<span data-ttu-id="f2725-p128">В примерах необязательных параметров параметр _callback_ указывается в качестве последнего параметра (после необязательных параметров, а также после объекта аргумента _Options_ ). Кроме того, можно указать параметр _обратного вызова_ в встроенном объекте JSON или в `options` объекте. Однако вы можете передать параметр _обратного вызова_ только в одном расположении: либо в объекте _Options_ (встроенном или созданном извне), либо в качестве последнего параметра, но не в обоих параметрах.</span><span class="sxs-lookup"><span data-stu-id="f2725-p128">In both optional parameter examples, the _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object). Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object. However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>


## <a name="see-also"></a><span data-ttu-id="f2725-221">См. также</span><span class="sxs-lookup"><span data-stu-id="f2725-221">See also</span></span>

- [<span data-ttu-id="f2725-222">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="f2725-222">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="f2725-223">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="f2725-223">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
