---
ms.date: 09/20/2018
description: Рекомендации и рекомендуемые шаблоны для настраиваемых функций Excel.
title: Рекомендации по настраиваемым функциям
ms.openlocfilehash: 3934910c397aea348c4fe2d7f95f1dc20ebeb4d3
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985790"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="a0b10-103">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="a0b10-103">Custom functions best practices</span></span>

<span data-ttu-id="a0b10-104">В этой статье описаны рекомендации по разработке настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="a0b10-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="a0b10-105">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="a0b10-105">Error handling</span></span>

<span data-ttu-id="a0b10-106">При построении надстройки, которая определяет настраиваемые функции, не забудьте включить логику обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="a0b10-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="a0b10-107">Обработка ошибок для настраиваемых функций такая же, как и в случае [обработки ошибок для API JavaScript Excel в целом](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="a0b10-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="a0b10-108">В следующем примере кода `.catch` будут обрабатываться все ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="a0b10-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="error-logging"></a><span data-ttu-id="a0b10-109">Ведение журнала ошибок</span><span class="sxs-lookup"><span data-stu-id="a0b10-109">Error logging</span></span>

<span data-ttu-id="a0b10-110">Можно включить журнал ведения  ошибки для настраиваемых функций надстройки несколькими способами, такими как:</span><span class="sxs-lookup"><span data-stu-id="a0b10-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="a0b10-111">Используйте [ведение журнала времени выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) для отладки XML-файла манифеста вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="a0b10-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="a0b10-112">Используйте `console.log` операторы в коде настраиваемых функций для отправки выходных данных в консоль в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="a0b10-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="a0b10-113">В настоящее время регистрация времени выполнения доступна только для классических приложений Office 2016.</span><span class="sxs-lookup"><span data-stu-id="a0b10-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="a0b10-114">Отладка</span><span class="sxs-lookup"><span data-stu-id="a0b10-114">Debugging</span></span>

<span data-ttu-id="a0b10-115">На данный момент наилучшим методом отладки пользовательских функций Excel является предварительная [загрузка неопубликованной надстройки](../testing/sideload-office-add-ins-for-testing.md) в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a0b10-115">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within Excel Online.</span></span> <span data-ttu-id="a0b10-116">Затем вы можете выполнить отладку настраиваемых функций с помощью [собственного средства отладки F12 вашего браузера](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="a0b10-116">Then you can debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span>

<span data-ttu-id="a0b10-117">Если надстройку не удалось зарегистрировать, [проверьте правильность настройки сертификатов SSL](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, где размещено приложение надстройки.</span><span class="sxs-lookup"><span data-stu-id="a0b10-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="a0b10-118">Сопоставление имен</span><span class="sxs-lookup"><span data-stu-id="a0b10-118">Mapping names</span></span>

<span data-ttu-id="a0b10-119">По умолчанию, имя настраиваемой функции в файл JavaScript обычно объявляется полностью с помощьюпрописных букв и в точности соответствует имени функции, которую видят конечные пользователи в Excel.</span><span class="sxs-lookup"><span data-stu-id="a0b10-119">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="a0b10-120">Тем не менее, можно изменить это с помощью `CustomFunctionsMappings` объекта для сопоставления одного или нескольких имен функции из файла JavaScript с разными значениями, которые  конечные пользователи увидят как имена функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="a0b10-120">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="a0b10-121">Эта функция полезна, если вы используете синтаксис методов uglifier, webpack или import - у каждого из которых есть трудности с именами функций в верхнем регистре.</span><span class="sxs-lookup"><span data-stu-id="a0b10-121">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="a0b10-122">`CustomFunctionsMappings` возможно не обязателен для проектов, использующих JavaScript, но этот объект необходимо использовать, если в вашем проекте применяется TypeScript.</span><span class="sxs-lookup"><span data-stu-id="a0b10-122">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="a0b10-123">В следующем примере кода определяется одна пара "ключ-значение", которая сопоставляет имя функции JavaScript `plusFortyTwo` с `ADD42` именем функции в пользовательском Интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="a0b10-123">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="a0b10-124">Когда конечный пользователь выбирает `ADD42` функцию в Excel, `plusFortyTwo`запускается функция JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0b10-124">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="a0b10-125">В следующем примере кода определяются две пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="a0b10-125">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="a0b10-126">Первая пара сопоставляет имя функции JavaScript `plusFifty` с `ADD50` именем функции в пользовательском Интерфейсе Excel и вторая пара сопоставляет имя функции JavaScript `plusOneHundred` с `ADD100` именем функции в пользовательском Интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="a0b10-126">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="a0b10-127">Когда конечный пользователь выбирает `ADD50` функцию в Excel, `plusFifty`запускается функция JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0b10-127">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="a0b10-128">Когда конечный пользователь выбирает `ADD100` функцию в Excel, `plusOneHundred`запускается функция JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0b10-128">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a><span data-ttu-id="a0b10-129">См. также</span><span class="sxs-lookup"><span data-stu-id="a0b10-129">See also</span></span>

* [<span data-ttu-id="a0b10-130">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="a0b10-130">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a0b10-131">Настраиваемые функции метаданных</span><span class="sxs-lookup"><span data-stu-id="a0b10-131">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a0b10-132">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="a0b10-132">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
