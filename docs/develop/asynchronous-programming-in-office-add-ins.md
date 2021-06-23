---
title: Асинхронное программирование в случае надстроек Office
description: Узнайте, как Office JavaScript использует асинхронное программирование в Office надстройки.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 42cf2d8e1b0d5185866a55152517683031da3b3d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076268"
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="a062b-103">Асинхронное программирование в надстройках для Office</span><span class="sxs-lookup"><span data-stu-id="a062b-103">Asynchronous programming in Office Add-ins</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="a062b-104">Почему в API Надстройки Office используется асинхронное программирование?</span><span class="sxs-lookup"><span data-stu-id="a062b-104">Why does the Office Add-ins API use asynchronous programming?</span></span> <span data-ttu-id="a062b-105">JavaScript — это язык однопотокового программирования, поэтому если скрипт вызывает продолжительный синхронный процесс, исполнение всех последующих скриптов будет заблокировано до завершения этого процесса.</span><span class="sxs-lookup"><span data-stu-id="a062b-105">Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes.</span></span> <span data-ttu-id="a062b-106">Поскольку некоторые операции Office веб-клиентов (но и богатых клиентов) могут блокировать выполнение, если они выполняются синхронно, большинство API JavaScript Office выполняться асинхронно.</span><span class="sxs-lookup"><span data-stu-id="a062b-106">Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the Office JavaScript APIs are designed to execute asynchronously.</span></span> <span data-ttu-id="a062b-107">Это позволяет Office надстройки быстро и быстро.</span><span class="sxs-lookup"><span data-stu-id="a062b-107">This makes sure that Office Add-ins are responsive and fast.</span></span> <span data-ttu-id="a062b-108">При работе с асинхронными методами зачастую требуется создавать функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a062b-108">It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="a062b-109">Имена всех асинхронных методов в API заканчиваются "Async", такими как `Document.getSelectedDataAsync` , `Binding.getDataAsync` или `Item.loadCustomPropertiesAsync` методы.</span><span class="sxs-lookup"><span data-stu-id="a062b-109">The names of all asynchronous methods in the API end with "Async", such as the `Document.getSelectedDataAsync`, `Binding.getDataAsync`, or `Item.loadCustomPropertiesAsync` methods.</span></span> <span data-ttu-id="a062b-110">При вызове асинхронного метода он выполняется немедленно и все дополнительные скрипты могут продолжать работу.</span><span class="sxs-lookup"><span data-stu-id="a062b-110">When an "Async" method is called, it executes immediately and any subsequent script execution can continue.</span></span> <span data-ttu-id="a062b-111">Необязательная функция обратного вызова, передаваемая в асинхронный метод, выполняется тогда, когда готовы данные или запрашиваемая операция.</span><span class="sxs-lookup"><span data-stu-id="a062b-111">The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready.</span></span> <span data-ttu-id="a062b-112">Обычно это происходит быстро, но иногда возможен возврат с небольшой задержкой.</span><span class="sxs-lookup"><span data-stu-id="a062b-112">This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="a062b-113">На следующей схеме показан поток выполнения для вызова метода "Async", который считывает данные, выбранные пользователем в документе, открытом на сервере Word или Excel.</span><span class="sxs-lookup"><span data-stu-id="a062b-113">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word or Excel.</span></span> <span data-ttu-id="a062b-114">В момент, когда выполняется вызов "Async", поток выполнения JavaScript может выполнять любую дополнительную клиентскую обработку (хотя ни один из них не отображается на схеме).</span><span class="sxs-lookup"><span data-stu-id="a062b-114">At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing (although none are shown in the diagram).</span></span> <span data-ttu-id="a062b-115">Когда метод "Async" возвращается, обратное вызов возобновляет выполнение в потоке, и надстройка может получить доступ к данным, сделать что-то с ним и отобразить результат.</span><span class="sxs-lookup"><span data-stu-id="a062b-115">When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result.</span></span> <span data-ttu-id="a062b-116">Тот же асинхронный шаблон выполнения сохраняется при работе с Office клиентских приложений, таких как Word 2013 или Excel 2013.</span><span class="sxs-lookup"><span data-stu-id="a062b-116">The same asynchronous execution pattern holds when working with the Office rich client applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="a062b-117">*Рис. 1. Процесс выполнения при асинхронном программировании*</span><span class="sxs-lookup"><span data-stu-id="a062b-117">*Figure 1. Asynchronous programming execution flow*</span></span>

![Схема, показывающая взаимодействие командного выполнения со временем с пользователем, страницей надстройки и сервером веб-приложения, на котором размещена надстройка.](../images/office-addins-asynchronous-programming-flow.png)

<span data-ttu-id="a062b-p104">Поддержка этой асинхронной конструкции как в полнофункциональных, так и в веб-клиентах предусмотрена в рамках стратегии проектирования "однократное написание — запуск на нескольких платформах" модели разработки надстроек Office. Например, вы можете создать надстройку области задач или контентную надстройку на единой базе кода, которая будет работать как в Excel 2013, так и в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a062b-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel on the web.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="a062b-121">Написание функции обратного вызова для асинхронного метода</span><span class="sxs-lookup"><span data-stu-id="a062b-121">Writing the callback function for an "Async" method</span></span>

<span data-ttu-id="a062b-122">Функция вызова, которую вы  передаете в качестве аргумента вызова в метод "Async", должна объявить один параметр, который будет использовать время выполнения надстройки для предоставления доступа к объекту [AsyncResult](/javascript/api/office/office.asyncresult) при выполнении функции вызова.</span><span class="sxs-lookup"><span data-stu-id="a062b-122">The callback function you pass as the _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](/javascript/api/office/office.asyncresult) object when the callback function executes.</span></span> <span data-ttu-id="a062b-123">Можно записать:</span><span class="sxs-lookup"><span data-stu-id="a062b-123">You can write:</span></span>

- <span data-ttu-id="a062b-124">Анонимная функция, которая должна быть написана и передана непосредственно в  соответствии с вызовом метода "Async" в качестве параметра вызова метода "Async".</span><span class="sxs-lookup"><span data-stu-id="a062b-124">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the _callback_ parameter of the "Async" method.</span></span>

- <span data-ttu-id="a062b-125">Именоваемая функция, передав  имя этой функции в качестве параметра вызова метода "Async".</span><span class="sxs-lookup"><span data-stu-id="a062b-125">A named function, passing the name of that function as the _callback_ parameter of an "Async" method.</span></span>

<span data-ttu-id="a062b-p106">Анонимную функцию удобно использовать, если код такой функции будет использован всего один раз (так как у нее нет имени, вы не сможете сослаться на нее в другой части кода). Именованные функции применяются, если необходимо многократно использовать функцию обратного вызова для нескольких асинхронных методов.</span><span class="sxs-lookup"><span data-stu-id="a062b-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>

### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="a062b-128">Написание анонимной функции обратного вызова</span><span class="sxs-lookup"><span data-stu-id="a062b-128">Writing an anonymous callback function</span></span>

<span data-ttu-id="a062b-129">Следующая функция анонимного обратного вызова объявляет один параметр с именем, который извлекает данные из свойства `result` [AsyncResult.value](/javascript/api/office/office.asyncresult#value) при возвращении обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a062b-129">The following anonymous callback function declares a single parameter named `result` that retrieves data from the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property when the callback returns.</span></span>

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="a062b-130">В следующем примере показано, как передать эту функцию анонимного вызова в строку в контексте полного вызова метода "Async" к `Document.getSelectedDataAsync` методу.</span><span class="sxs-lookup"><span data-stu-id="a062b-130">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the `Document.getSelectedDataAsync` method.</span></span>

- <span data-ttu-id="a062b-131">Первый аргумент _coercionType_ указывает, чтобы вернуть выбранные данные `Office.CoercionType.Text` в виде строки текста.</span><span class="sxs-lookup"><span data-stu-id="a062b-131">The first _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>

- <span data-ttu-id="a062b-132">Второй аргумент _вызова_ — анонимная функция, переданная методу в строке.</span><span class="sxs-lookup"><span data-stu-id="a062b-132">The second _callback_ argument is the anonymous function passed in-line to the method.</span></span> <span data-ttu-id="a062b-133">При выполнении функции для  отображения данных, выбранных пользователем в документе, используется параметр результатов для доступа к свойству `value` `AsyncResult` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-133">When the function executes, it uses the _result_ parameter to access the `value` property of the `AsyncResult` object to display the data selected by the user in the document.</span></span>

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

<span data-ttu-id="a062b-134">Вы также можете использовать параметр функции вызова для доступа к другим свойствам `AsyncResult` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-134">You can also use the parameter of your callback function to access other properties of the `AsyncResult` object.</span></span> <span data-ttu-id="a062b-135">Используйте свойство [AsyncResult.status](/javascript/api/office/office.asyncresult#status), чтобы определить, успешно ли был выполнен вызов.</span><span class="sxs-lookup"><span data-stu-id="a062b-135">Use the [AsyncResult.status](/javascript/api/office/office.asyncresult#status) property to determine if the call succeeded or failed.</span></span> <span data-ttu-id="a062b-136">Если не удалось выполнить вызов, можно использовать свойство [AsyncResult.error](/javascript/api/office/office.asyncresult#error), чтобы получить доступ к объекту [Error](/javascript/api/office/office.error) и получить сведения об ошибке.</span><span class="sxs-lookup"><span data-stu-id="a062b-136">If your call fails you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#error) property to access an [Error](/javascript/api/office/office.error) object for error information.</span></span>

<span data-ttu-id="a062b-137">Дополнительные сведения об использовании метода см. в публикации `getSelectedDataAsync` Read and write data to the active selection in a document or [spreadsheet.](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)</span><span class="sxs-lookup"><span data-stu-id="a062b-137">For more information about using the `getSelectedDataAsync` method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 

### <a name="writing-a-named-callback-function"></a><span data-ttu-id="a062b-138">Написание именованной функции обратного вызова</span><span class="sxs-lookup"><span data-stu-id="a062b-138">Writing a named callback function</span></span>

<span data-ttu-id="a062b-139">Кроме того, можно написать именоваемую  функцию и передать ее имя параметру вызова метода "Async".</span><span class="sxs-lookup"><span data-stu-id="a062b-139">Alternatively, you can write a named function and pass its name to the _callback_ parameter of an "Async" method.</span></span> <span data-ttu-id="a062b-140">Например, предыдущий пример можно изменить так, чтобы передавать функцию с именем `writeDataCallback` в качестве параметра _callback_.</span><span class="sxs-lookup"><span data-stu-id="a062b-140">For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>

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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="a062b-141">Что возвращается в свойство AsyncResult.value?</span><span class="sxs-lookup"><span data-stu-id="a062b-141">Differences in what's returned to the AsyncResult.value property</span></span>

<span data-ttu-id="a062b-142">Свойства объекта возвращают те же виды информации в функцию обратного вызова, переданную всем методам `asyncContext` `status` `error` `AsyncResult` async.</span><span class="sxs-lookup"><span data-stu-id="a062b-142">The `asyncContext`, `status`, and `error` properties of the `AsyncResult` object return the same kinds of information to the callback function passed to all "Async" methods.</span></span> <span data-ttu-id="a062b-143">Однако то, что возвращается в свойство, зависит от функциональности метода `AsyncResult.value` "Async".</span><span class="sxs-lookup"><span data-stu-id="a062b-143">However, what's returned to the `AsyncResult.value` property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="a062b-144">Например, методы `addHandlerAsync` (binding, [](/javascript/api/office/office.binding) [CustomXmlPart,](/javascript/api/office/office.customxmlpart) [Document,](/javascript/api/office/office.document) [RoamingSettings](/javascript/api/outlook/office.roamingsettings)и [Параметры](/javascript/api/office/office.settings) объектов) используются для добавления функций обработки событий в элементы, представленные этими объектами.</span><span class="sxs-lookup"><span data-stu-id="a062b-144">For example, the `addHandlerAsync` methods (of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [Settings](/javascript/api/office/office.settings) objects) are used to add event handler functions to the items represented by these objects.</span></span> <span data-ttu-id="a062b-145">Вы можете получить доступ к свойству из функции обратного вызова, передаемой любому из методов, но так как данные или объект не доступны при добавлении обработчицы событий, свойство всегда возвращается неопределенным, если вы попытаетесь получить к нему `AsyncResult.value` `addHandlerAsync` `value` доступ. </span><span class="sxs-lookup"><span data-stu-id="a062b-145">You can access the `AsyncResult.value` property from the callback function you pass to any of the `addHandlerAsync` methods, but since no data or object is being accessed when you add an event handler, the `value` property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="a062b-146">С другой стороны, при вызове метода он возвращает данные, выбранные пользователем в документе, в свойство `Document.getSelectedDataAsync` `AsyncResult.value` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a062b-146">On the other hand, if you call the `Document.getSelectedDataAsync` method, it returns the data the user selected in the document to the `AsyncResult.value` property in the callback.</span></span> <span data-ttu-id="a062b-147">Или, если вы называете метод [Bindings.getAllAsync,](/javascript/api/office/office.bindings#getallasync-options--callback-) он возвращает массив всех объектов `Binding` в документе.</span><span class="sxs-lookup"><span data-stu-id="a062b-147">Or, if you call the [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) method, it returns an array of all of the `Binding` objects in the document.</span></span> <span data-ttu-id="a062b-148">И если вы назовете метод [Bindings.getByIdAsync,](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) он возвращает один `Binding` объект.</span><span class="sxs-lookup"><span data-stu-id="a062b-148">And, if you call the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method, it returns a single `Binding` object.</span></span>

<span data-ttu-id="a062b-149">Описание того, что возвращается в свойство для метода, см. в разделе "Значение обратного вызова" справочной `AsyncResult.value` `Async` темы этого метода.</span><span class="sxs-lookup"><span data-stu-id="a062b-149">For a description of what's returned to the `AsyncResult.value` property for an `Async` method, see the "Callback value" section of that method's reference topic.</span></span> <span data-ttu-id="a062b-150">Сводку всех объектов, которые предоставляют методы, см. в таблице в нижней части темы `Async` [объекта AsyncResult.](/javascript/api/office/office.asyncresult)</span><span class="sxs-lookup"><span data-stu-id="a062b-150">For a summary of all of the objects that provide `Async` methods, see the table at the bottom of the [AsyncResult](/javascript/api/office/office.asyncresult) object topic.</span></span>

## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="a062b-151">Шаблоны асинхронного программирования</span><span class="sxs-lookup"><span data-stu-id="a062b-151">Asynchronous programming patterns</span></span>

<span data-ttu-id="a062b-152">API Office JavaScript поддерживает два вида асинхронных шаблонов программирования:</span><span class="sxs-lookup"><span data-stu-id="a062b-152">The Office JavaScript API supports two kinds of asynchronous programming patterns:</span></span>

- <span data-ttu-id="a062b-153">С использованием вложенных обратных вызовов</span><span class="sxs-lookup"><span data-stu-id="a062b-153">Using nested callbacks</span></span>
- <span data-ttu-id="a062b-154">С использованием шаблона promise</span><span class="sxs-lookup"><span data-stu-id="a062b-154">Using the promises pattern</span></span>

<span data-ttu-id="a062b-p114">При асинхронном программировании с использованием функций обратного вызова зачастую требуется вкладывать возвращаемый результат одного обратного вызова в один или несколько других обратных вызовов. В этом случае вы можете использовать вложенные обратные вызовы асинхронных методов API.</span><span class="sxs-lookup"><span data-stu-id="a062b-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="a062b-157">Использование вложенных обратных вызовов — это шаблон программирования, знакомый большинству разработчиков на языке JavaScript, но код с глубоко вложенными обратными вызовами может быть труден для чтения и понимания.</span><span class="sxs-lookup"><span data-stu-id="a062b-157">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand.</span></span> <span data-ttu-id="a062b-158">В качестве альтернативы вложенным вызовам Office API JavaScript также поддерживает реализацию шаблона обещаний.</span><span class="sxs-lookup"><span data-stu-id="a062b-158">As an alternative to nested callbacks, the Office JavaScript API also supports an implementation of the promises pattern.</span></span>

> [!NOTE]
> <span data-ttu-id="a062b-159">В текущей версии API Office JavaScript  встроенная поддержка шаблона обещаний работает только с кодом для привязки в Excel таблицах и документах [Word](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="a062b-159">In the current version of the Office JavaScript API, *built-in* support for the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span> <span data-ttu-id="a062b-160">Однако вы можете обернуть другие функции, которые имеют обратное вызовы в вашей собственной настраиваемой функции возврата обещаний.</span><span class="sxs-lookup"><span data-stu-id="a062b-160">However, you can wrap other functions that have callbacks inside your own custom Promise-returning function.</span></span> <span data-ttu-id="a062b-161">Дополнительные сведения см. в ссылке [Wrap Common API in Promise-returning functions.](#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="a062b-161">For more information, see [Wrap Common APIs in Promise-returning functions](#wrap-common-apis-in-promise-returning-functions).</span></span>

### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="a062b-162">Асинхронное программирование с использованием вложенных функций обратного вызова</span><span class="sxs-lookup"><span data-stu-id="a062b-162">Asynchronous programming using nested callback functions</span></span>

<span data-ttu-id="a062b-p117">Зачастую для какой-либо задачи необходимо выполнять несколько асинхронных операций. Для этого можно вкладывать один асинхронный вызов в другой.</span><span class="sxs-lookup"><span data-stu-id="a062b-p117">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span>

<span data-ttu-id="a062b-165">В следующем примере кода показано, как вложить два асинхронных вызова.</span><span class="sxs-lookup"><span data-stu-id="a062b-165">The following code example nests two asynchronous calls.</span></span>

- <span data-ttu-id="a062b-166">Сначала вызывается метод [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) для получения доступа к привязке в документе с именем "MyBinding".</span><span class="sxs-lookup"><span data-stu-id="a062b-166">First, the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method is called to access a binding in the document named "MyBinding".</span></span> <span data-ttu-id="a062b-167">Объект, `AsyncResult` возвращающийся к параметру обратного вызова, предоставляет доступ к указанному объекту привязки `result` из `AsyncResult.value` свойства.</span><span class="sxs-lookup"><span data-stu-id="a062b-167">The `AsyncResult` object returned to the `result` parameter of that callback provides access to the specified binding object from the `AsyncResult.value` property.</span></span>
- <span data-ttu-id="a062b-168">Затем объект привязки, доступный из первого параметра, используется для вызова метода `result` [Binding.getDataAsync.](/javascript/api/office/office.binding#getdataasync-options--callback-)</span><span class="sxs-lookup"><span data-stu-id="a062b-168">Then, the binding object accessed from the first `result` parameter is used to call the [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) method.</span></span>
- <span data-ttu-id="a062b-169">Наконец, параметр вызова, переданного методу, используется для отображения `result2` `Binding.getDataAsync` данных в привязке.</span><span class="sxs-lookup"><span data-stu-id="a062b-169">Finally, the `result2` parameter of the callback passed to the `Binding.getDataAsync` method is used to display the data in the binding.</span></span>

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

<span data-ttu-id="a062b-170">Этот базовый шаблон вложенного вызова можно использовать для всех асинхронных методов в Office API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a062b-170">This basic nested callback pattern can be used for all asynchronous methods in the Office JavaScript API.</span></span>

<span data-ttu-id="a062b-171">В следующих разделах показано, как использовать анонимные или именованные функции для вложенных обратных вызовов в асинхронных методах.</span><span class="sxs-lookup"><span data-stu-id="a062b-171">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>

#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="a062b-172">Использование анонимных функций для вложенных обратных вызовов</span><span class="sxs-lookup"><span data-stu-id="a062b-172">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="a062b-173">В следующем примере две анонимные функции объявляются inline и передаются в вложенные обратное вызовы и `getByIdAsync` `getDataAsync` методы.</span><span class="sxs-lookup"><span data-stu-id="a062b-173">In the following example, two anonymous functions are declared inline and passed into the `getByIdAsync` and `getDataAsync` methods as nested callbacks.</span></span> <span data-ttu-id="a062b-174">Поскольку это простые и встроенные функции, их назначение сразу же становится понятным.</span><span class="sxs-lookup"><span data-stu-id="a062b-174">Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>

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

#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="a062b-175">Использование именованных функций для вложенных обратных вызовов</span><span class="sxs-lookup"><span data-stu-id="a062b-175">Using named functions for nested callbacks</span></span>

<span data-ttu-id="a062b-176">В сложных реализациях может оказаться полезным использовать именованные функции для упрощения чтения, поддержки и повторного использования.</span><span class="sxs-lookup"><span data-stu-id="a062b-176">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse.</span></span> <span data-ttu-id="a062b-177">В следующем примере две анонимные функции из примера в предыдущем разделе были переписаны как функции с `deleteAllData` именем и `showResult` .</span><span class="sxs-lookup"><span data-stu-id="a062b-177">In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named `deleteAllData` and `showResult`.</span></span> <span data-ttu-id="a062b-178">Эти названные функции затем передаются в методы обратного вызова по `getByIdAsync` `deleteAllDataValuesAsync` имени.</span><span class="sxs-lookup"><span data-stu-id="a062b-178">These named functions are then passed into the `getByIdAsync` and `deleteAllDataValuesAsync` methods as callbacks by name.</span></span>

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

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="a062b-179">Асинхронное программирование с применением шаблона, предусматривающего использование обещаний для получения доступа к данным в привязках</span><span class="sxs-lookup"><span data-stu-id="a062b-179">Asynchronous programming using the promises pattern to access data in bindings</span></span>

<span data-ttu-id="a062b-p121">Если применяется шаблон программирования, предусматривающий использование обещаний, в коде не нужно указывать передачу функции обратного вызова и ожидание ее возвращения для продолжения выполнения. В этом случае сразу возвращается объект обещания, который представляет нужный результат. Но в отличие от традиционного синхронного программирования, в этом случае получение обещанного результата на самом деле откладывается до тех пор, пока среда выполнения надстроек Office не сможет выполнить запрос. Обработчик _onError_ предоставляется для ситуаций, когда запрос не может быть выполнен.</span><span class="sxs-lookup"><span data-stu-id="a062b-p121">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>

<span data-ttu-id="a062b-183">API Office JavaScript предоставляет [метод Office.select](/javascript/api/office#office-select-expression--callback-) для поддержки шаблона обещаний для работы с существующими объектами привязки.</span><span class="sxs-lookup"><span data-stu-id="a062b-183">The Office JavaScript API provides the [Office.select](/javascript/api/office#office-select-expression--callback-) method to support the promises pattern for working with existing binding objects.</span></span> <span data-ttu-id="a062b-184">Объект promise, возвращенный методу, поддерживает только четыре метода, к которые можно получить доступ непосредственно из объекта `Office.select` [Binding:](/javascript/api/office/office.binding) [getDataAsync,](/javascript/api/office/office.binding#getdataasync-options--callback-) [setDataAsync,](/javascript/api/office/office.binding#setdataasync-data--options--callback-) [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)и [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="a062b-184">The promise object returned to the `Office.select` method supports only the four methods that you can access directly from the [Binding](/javascript/api/office/office.binding) object: [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-), and [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span></span>

<span data-ttu-id="a062b-185">Шаблон promise для работы с привязками принимает такую форму:</span><span class="sxs-lookup"><span data-stu-id="a062b-185">The promises pattern for working with bindings takes this form:</span></span>

<span data-ttu-id="a062b-186">**Office.select**_(selectorExpression_, _onError_**).** _BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="a062b-186">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="a062b-187">Параметр _selectorExpression_ принимает форму, в которой bindingId — это имя () привязки, созданной ранее в документе или таблице (с помощью одного из методов `"bindings#bindingId"`  `id` addFrom `Bindings` коллекции: `addFromNamedItemAsync` , или `addFromPromptAsync` `addFromSelectionAsync` ).</span><span class="sxs-lookup"><span data-stu-id="a062b-187">The _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where _bindingId_ is the name ( `id`) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the `Bindings` collection: `addFromNamedItemAsync`, `addFromPromptAsync`, or `addFromSelectionAsync`).</span></span> <span data-ttu-id="a062b-188">Например, выражение селектора указывает, что необходимо получить доступ к привязке с `bindings#cities` **id** "cities".</span><span class="sxs-lookup"><span data-stu-id="a062b-188">For example, the selector expression `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="a062b-189">Параметр _onError_ — это функция обработки ошибок, которая принимает один параметр типа, который может использоваться для доступа к объекту, если метод не может получить доступ к указанной `AsyncResult` `Error` `select` привязке.</span><span class="sxs-lookup"><span data-stu-id="a062b-189">The _onError_ parameter is an error handling function which takes a single parameter of type `AsyncResult` that can be used to access an `Error` object, if the `select` method fails to access the specified binding.</span></span> <span data-ttu-id="a062b-190">В следующем примере показана базовая функция обработки ошибки, которую можно передать в параметр _onError_.</span><span class="sxs-lookup"><span data-stu-id="a062b-190">The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>

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

<span data-ttu-id="a062b-191">Замените местообладатель _BindingObjectAsyncMethod_ вызовом на любой из четырех методов объекта, поддерживаемых объектом `Binding` promise: `getDataAsync` , , , `setDataAsync` или `addHandlerAsync` `removeHandlerAsync` .</span><span class="sxs-lookup"><span data-stu-id="a062b-191">Replace the _BindingObjectAsyncMethod_ placeholder with a call to any of the four `Binding` object methods supported by the promise object: `getDataAsync`, `setDataAsync`, `addHandlerAsync`, or `removeHandlerAsync`.</span></span> <span data-ttu-id="a062b-192">Вызовы этих методов не поддерживают дополнительные шаблоны promise.</span><span class="sxs-lookup"><span data-stu-id="a062b-192">Calls to these methods don't support additional promises.</span></span> <span data-ttu-id="a062b-193">Их нужно вызывать с помощью [шаблона функции вложенного обратного вызова](#asynchronous-programming-using-nested-callback-functions).</span><span class="sxs-lookup"><span data-stu-id="a062b-193">You must call them using the [nested callback function pattern](#asynchronous-programming-using-nested-callback-functions).</span></span>

<span data-ttu-id="a062b-194">После выполнения обещания объекта его можно повторно использовать в цепном вызове метода, как при привязке (время выполнения надстройки не будет асинхронно выполнять `Binding` обещание).</span><span class="sxs-lookup"><span data-stu-id="a062b-194">After a `Binding` object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise).</span></span> <span data-ttu-id="a062b-195">Если обещание объекта не может быть выполнено, время выполнения надстройки снова будет пытаться получить доступ к объекту привязки при следующем вызове одного из его асинхронных `Binding` методов.</span><span class="sxs-lookup"><span data-stu-id="a062b-195">If the `Binding` object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="a062b-196">В следующем примере кода используется метод для получения привязки с "из коллекции", а затем вызывает метод `select` `id` `cities` `Bindings` [addHandlerAsync,](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) чтобы добавить обработник событий для события [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) привязки.</span><span class="sxs-lookup"><span data-stu-id="a062b-196">The following code example uses the `select` method to retrieve a binding with the `id` "`cities`" from the `Bindings` collection, and then calls the [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event of the binding.</span></span>

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> <span data-ttu-id="a062b-197">Обещание `Binding` объекта, возвращаемого методом, предоставляет доступ только `Office.select` к четырем методам `Binding` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-197">The `Binding` object promise returned by the `Office.select` method provides access to only the four methods of the `Binding` object.</span></span> <span data-ttu-id="a062b-198">Если вам нужно получить доступ к любому из других участников объекта, вместо этого необходимо использовать свойство и методы для `Binding` `Document.bindings` получения `Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-198">If you need to access any of the other members of the `Binding` object, instead you must use the `Document.bindings` property and `Bindings.getByIdAsync` or `Bindings.getAllAsync` methods to retrieve the `Binding` object.</span></span> <span data-ttu-id="a062b-199">Например, если вам необходимо получить доступ к любым свойствам объекта (свойствам или свойствам) или получить доступ к свойствам объектов `Binding` `document` `id` `type` [MatrixBinding или TableBinding,](/javascript/api/office/office.matrixbinding) [](/javascript/api/office/office.tablebinding) `getByIdAsync` `getAllAsync` необходимо использовать или методы для получения `Binding` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-199">For example, if you need to access any of the `Binding` object's properties (the `document`, `id`, or `type` properties), or need to access the properties of the [MatrixBinding](/javascript/api/office/office.matrixbinding) or [TableBinding](/javascript/api/office/office.tablebinding) objects, you must use the `getByIdAsync` or `getAllAsync` methods to retrieve a `Binding` object.</span></span>

## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="a062b-200">Передача дополнительных параметров в асинхронные методы</span><span class="sxs-lookup"><span data-stu-id="a062b-200">Passing optional parameters to asynchronous methods</span></span>

<span data-ttu-id="a062b-201">Общий синтаксис методов "Async" следует следующему шаблону:</span><span class="sxs-lookup"><span data-stu-id="a062b-201">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="a062b-202">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`</span><span class="sxs-lookup"><span data-stu-id="a062b-202">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="a062b-p128">Все асинхронные методы поддерживают дополнительные параметры, которые передаются в виде объекта JSON, содержащего один или несколько дополнительных параметров. Объект JSON, содержащий дополнительные параметры, является неупорядоченной коллекцией пар "ключ-значение" с разделителем ":". Каждая пара в объекте разделяется точкой с запятой, а весь набор пар заключен в скобки. Ключом является имя параметра, а значением — значение, которое следует передать этому параметру.</span><span class="sxs-lookup"><span data-stu-id="a062b-p128">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="a062b-207">Вы можете создать объект JSON, содержащий дополнительные параметры в линию, или путем создания объекта и передачи его в качестве `options` _параметра параметра параметра параметра._</span><span class="sxs-lookup"><span data-stu-id="a062b-207">You can create the JSON object that contains optional parameters inline, or by creating an `options` object and passing that in as the _options_ parameter.</span></span>

### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="a062b-208">Передача дополнительных параметров в качестве встроенных</span><span class="sxs-lookup"><span data-stu-id="a062b-208">Passing optional parameters inline</span></span>

<span data-ttu-id="a062b-209">Например, синтаксис вызова метода [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) с необязательными параметрами в качестве встроенных выглядит так:</span><span class="sxs-lookup"><span data-stu-id="a062b-209">For example, the syntax for calling the [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

<span data-ttu-id="a062b-210">В этой форме синтаксиса вызовов два необязательных параметра, _coercionType_ и _asyncContext,_ определяются как объект JSON, закрытый в скобки.</span><span class="sxs-lookup"><span data-stu-id="a062b-210">In this form of the calling syntax, the two optional parameters, _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="a062b-211">В следующем примере показано, как вызвать метод, указав дополнительные `Document.setSelectedDataAsync` параметры в линию.</span><span class="sxs-lookup"><span data-stu-id="a062b-211">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters inline.</span></span>

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
> <span data-ttu-id="a062b-212">Дополнительные параметры можно задавать в объекте JSON в любом порядке, если их имена указываются правильно.</span><span class="sxs-lookup"><span data-stu-id="a062b-212">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>

### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="a062b-213">Передача дополнительных параметров в объекте options</span><span class="sxs-lookup"><span data-stu-id="a062b-213">Passing optional parameters in an options object</span></span>

<span data-ttu-id="a062b-214">Кроме того, можно создать объект с именем, который указывает необязательные параметры отдельно от вызова метода, а затем передать объект в `options` `options` качестве _аргумента параметра._</span><span class="sxs-lookup"><span data-stu-id="a062b-214">Alternatively, you can create an object named `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="a062b-215">В следующем примере показан один из способов создания объекта, в котором ( и т. д.) являются задатчиками фактических имен `options` `parameter1` и `value1` значений параметров.</span><span class="sxs-lookup"><span data-stu-id="a062b-215">The following example shows one way of creating the `options` object, where `parameter1`, `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="a062b-216">Когда указываются параметры [ValueFormat](/javascript/api/office/office.valueformat) и [FilterType](/javascript/api/office/office.filtertype), код будет таким:</span><span class="sxs-lookup"><span data-stu-id="a062b-216">Which looks like the following example when used to specify the [ValueFormat](/javascript/api/office/office.valueformat) and [FilterType](/javascript/api/office/office.filtertype) parameters.</span></span>

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="a062b-217">Вот еще один способ создания `options` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-217">Here's another way of creating the `options` object.</span></span>

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="a062b-218">Это выглядит как следующий пример, когда используется для указания параметров и `ValueFormat` `FilterType` параметров:</span><span class="sxs-lookup"><span data-stu-id="a062b-218">Which looks like the following example when used to specify the `ValueFormat` and `FilterType` parameters:</span></span>

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> <span data-ttu-id="a062b-219">При использовании любого метода создания объекта можно указать необязательные параметры в любом порядке, если их имена указаны `options` правильно.</span><span class="sxs-lookup"><span data-stu-id="a062b-219">When using either method of creating the `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="a062b-220">В следующем примере показано, как вызвать метод, указав необязательные `Document.setSelectedDataAsync` параметры `options` объекта.</span><span class="sxs-lookup"><span data-stu-id="a062b-220">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters in an `options` object.</span></span>

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

<span data-ttu-id="a062b-221">В обоих необязательных примерах параметров параметр _callback_ указывается как последний параметр (следуя за необязательными параметрами, или следуя объекту _аргумента параметра)._</span><span class="sxs-lookup"><span data-stu-id="a062b-221">In both optional parameter examples, the _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object).</span></span> <span data-ttu-id="a062b-222">Кроме того, параметр _callback_ можно указать либо во встроенном объекте JSON, либо в объекте `options`.</span><span class="sxs-lookup"><span data-stu-id="a062b-222">Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object.</span></span> <span data-ttu-id="a062b-223">Однако параметр _callback_ можно передать только одним из способов: или в объекте _options_ (встроенном или созданном внешне), или в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="a062b-223">However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>

## <a name="wrap-common-apis-in-promise-returning-functions"></a><span data-ttu-id="a062b-224">Оберните общие API в функциях возврата обещаний</span><span class="sxs-lookup"><span data-stu-id="a062b-224">Wrap Common APIs in Promise-returning functions</span></span>

<span data-ttu-id="a062b-225">Общие методы API (Outlook API) не возвращают [обещания.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)</span><span class="sxs-lookup"><span data-stu-id="a062b-225">The Common API (and Outlook API) methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="a062b-226">Поэтому нельзя использовать [](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) ожидание для приостановки выполнения до завершения асинхронной операции.</span><span class="sxs-lookup"><span data-stu-id="a062b-226">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="a062b-227">Если вам нужно `await` поведение, можно завернуть вызов метода в явно созданное обещание.</span><span class="sxs-lookup"><span data-stu-id="a062b-227">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span> 

<span data-ttu-id="a062b-228">Основной шаблон заключается в создании асинхронного метода, который  возвращает объект Promise немедленно и устраняет  объект Promise, когда внутренний метод завершается, или отклоняет объект, если метод не удается.</span><span class="sxs-lookup"><span data-stu-id="a062b-228">The basic pattern is to create an asynchronous method that returns a Promise object immediately and *resolves* that Promise object when the inner method completes, or *rejects* the object if the method fails.</span></span> <span data-ttu-id="a062b-229">Ниже приводится простой пример</span><span class="sxs-lookup"><span data-stu-id="a062b-229">The following is a simple example</span></span>

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

<span data-ttu-id="a062b-230">Когда этот метод необходимо ожидать, его можно назвать либо ключевым словом, либо функцией, передаемой `await` `then` функции.</span><span class="sxs-lookup"><span data-stu-id="a062b-230">When this method needs to be awaited, it can be called either with the `await` keyword or as the function passed to a `then` function.</span></span>

> [!NOTE]
> <span data-ttu-id="a062b-231">Этот метод особенно полезен при вызове одного из общих API внутри вызова метода в одной из моделей объектов, определенных `run` приложениям.</span><span class="sxs-lookup"><span data-stu-id="a062b-231">This technique is especially useful when you need to call one of the Common APIs inside a call of the `run` method in one of the application-specific object models.</span></span> <span data-ttu-id="a062b-232">Пример вышеуказанной функции см. вHome.js в [примере Word-Add-in-JavaScript-MDConversion.](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)</span><span class="sxs-lookup"><span data-stu-id="a062b-232">For an example of the function above being used in this way, see the file [Home.js in the sample Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).</span></span>

<span data-ttu-id="a062b-233">Ниже приводится пример с помощью TypeScript.</span><span class="sxs-lookup"><span data-stu-id="a062b-233">The following is an example using TypeScript.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="a062b-234">См. также</span><span class="sxs-lookup"><span data-stu-id="a062b-234">See also</span></span>

- [<span data-ttu-id="a062b-235">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="a062b-235">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="a062b-236">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="a062b-236">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
