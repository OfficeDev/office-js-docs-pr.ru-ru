---
ms.date: 09/20/2018
description: Рекомендации и рекомендуемые шаблоны для настраиваемых функций Excel.
title: Рекомендации по настраиваемым функциям
ms.openlocfilehash: 4fe0ddc36ce1b08ea360bb556121e76cd57c3823
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004912"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="696db-103">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="696db-103">Custom functions best practices</span></span>

<span data-ttu-id="696db-104">В этой статье описаны рекомендации по разработке настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="696db-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="696db-105">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="696db-105">Error handling</span></span>

<span data-ttu-id="696db-106">При построении надстройки, которая определяет настраиваемые функции, не забудьте включить логику обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="696db-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="696db-107">Обработка ошибок для настраиваемых функций совпадает с [обработкой ошибок для Excel API JavaScript в целом](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="696db-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="696db-108">В следующем примере кода метод `.catch` будет обрабатывать все ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="696db-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi.com/comments/" + x; 
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

## <a name="debugging"></a><span data-ttu-id="696db-109">Отладка</span><span class="sxs-lookup"><span data-stu-id="696db-109">Debugging</span></span>
<span data-ttu-id="696db-110">На данный момент наилучшим методом отладки пользовательских функций Excel является предварительная [загрузка неопубликованной надстройки](../testing/sideload-office-add-ins-for-testing.md) в **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="696db-110">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="696db-111">Затем вы можете выполнить отладку настраиваемых функций с помощью [собственного средства отладки F12 вашего веб-обозревателя](../testing/debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="696db-111">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span> <span data-ttu-id="696db-112">Используйте `console.log` операторы в коде настраиваемых функций для отправки выходных данных в консоль в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="696db-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

<span data-ttu-id="696db-113">Если надстройку не удалось зарегистрировать, [проверьте правильность настройки сертификатов SSL](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, где размещено приложение надстройки.</span><span class="sxs-lookup"><span data-stu-id="696db-113">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="696db-114">При тестировании надстройки в классическом приложении Office 2016 можно включить [регистрацию времени выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) для отладки проблем, связанных с XML-файлом манифеста вашей надстройки, а также несколько условий установки и выполнения.</span><span class="sxs-lookup"><span data-stu-id="696db-114">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span> 


## <a name="mapping-names"></a><span data-ttu-id="696db-115">Сопоставление имен</span><span class="sxs-lookup"><span data-stu-id="696db-115">Mapping names</span></span>

<span data-ttu-id="696db-116">По умолчанию, имя настраиваемой функции в файл JavaScript обычно объявляется полностью с помощьюпрописных букв и в точности соответствует имени функции, которую видят конечные пользователи в Excel.</span><span class="sxs-lookup"><span data-stu-id="696db-116">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="696db-117">Тем не менее, можно изменить это с помощью `CustomFunctionsMappings` объекта для сопоставления одного или нескольких имен функции из файла JavaScript с разными значениями, которые  конечные пользователи увидят как имена функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="696db-117">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="696db-118">Эта функция полезна, если вы используете синтаксис методов uglifier, webpack или import, у каждого из которых есть трудности с именами функций в верхнем регистре.</span><span class="sxs-lookup"><span data-stu-id="696db-118">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="696db-119">`CustomFunctionsMappings` может быть не обязательным для проектов, использующих JavaScript, но этот объект необходимо использовать, если в вашем проекте применяется TypeScript.</span><span class="sxs-lookup"><span data-stu-id="696db-119">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="696db-120">В следующем примере кода определяется одна пара "ключ-значение", которая сопоставляет имя функции JavaScript `plusFortyTwo` с `ADD42` именем функции в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="696db-120">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="696db-121">Когда конечный пользователь выбирает `ADD42` функцию в Excel, `plusFortyTwo`запускается функция JavaScript.</span><span class="sxs-lookup"><span data-stu-id="696db-121">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="696db-122">В следующем примере кода определяются две пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="696db-122">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="696db-123">Первая пара сопоставляет имя функции JavaScript `plusFifty` с `ADD50` именем функции в пользовательском Интерфейсе Excel и вторая пара сопоставляет имя функции JavaScript `plusOneHundred` с `ADD100` именем функции в пользовательском Интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="696db-123">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="696db-124">Когда конечный пользователь выбирает `ADD50` функцию в Excel, `plusFifty`запускается функция JavaScript.</span><span class="sxs-lookup"><span data-stu-id="696db-124">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="696db-125">Когда конечный пользователь выбирает `ADD100` функцию в Excel, `plusOneHundred`запускается функция JavaScript.</span><span class="sxs-lookup"><span data-stu-id="696db-125">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="696db-126">См. также</span><span class="sxs-lookup"><span data-stu-id="696db-126">See also</span></span>

- [<span data-ttu-id="696db-127">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="696db-127">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="696db-128">Настраиваемые функции метаданных</span><span class="sxs-lookup"><span data-stu-id="696db-128">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="696db-129">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="696db-129">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
