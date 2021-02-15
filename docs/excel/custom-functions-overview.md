---
ms.date: 01/08/2020
description: Создайте пользовательскую функцию Excel для надстройки Office
title: Создание пользовательских функций в Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 804895f3e10cac849dc20b67625e4f30164eb41d
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237674"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="69e08-103">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="69e08-103">Create custom functions in Excel</span></span>

<span data-ttu-id="69e08-104">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="69e08-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="69e08-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="69e08-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="69e08-106">Ниже на анимированном изображении показано, как рабочая книга вызывает функцию, созданную вами с помощью JavaScript или Typescript.</span><span class="sxs-lookup"><span data-stu-id="69e08-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="69e08-107">В этом примере пользовательская функция `=MYFUNCTION.SPHEREVOLUME` рассчитывает объем сферы.</span><span class="sxs-lookup"><span data-stu-id="69e08-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="69e08-108">Приведенный ниже код определяет пользовательскую функцию `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="69e08-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!TIP]
> <span data-ttu-id="69e08-109">Если надстройка пользовательской функции использует область задач или кнопку ленты (помимо выполнения кода пользовательской функции), вам потребуется настроить общую среду выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="69e08-109">If your custom function add-in will use a task pane or a ribbon button, in addition to running custom function code, you will need to set up a shared JavaScript runtime.</span></span> <span data-ttu-id="69e08-110">Дополнительные сведения см. в статье [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="69e08-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="69e08-111">Как определена пользовательская функция в коде</span><span class="sxs-lookup"><span data-stu-id="69e08-111">How a custom function is defined in code</span></span>

<span data-ttu-id="69e08-112">Если использовать [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, он создаст файлы, управляющие вашими функциями и областью задач.</span><span class="sxs-lookup"><span data-stu-id="69e08-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="69e08-113">Мы сосредоточимся на файлах, которые важны для пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="69e08-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="69e08-114">Файл</span><span class="sxs-lookup"><span data-stu-id="69e08-114">File</span></span> | <span data-ttu-id="69e08-115">Формат файла</span><span class="sxs-lookup"><span data-stu-id="69e08-115">File format</span></span> | <span data-ttu-id="69e08-116">Описание</span><span class="sxs-lookup"><span data-stu-id="69e08-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="69e08-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="69e08-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="69e08-118">или</span><span class="sxs-lookup"><span data-stu-id="69e08-118">or</span></span><br/><span data-ttu-id="69e08-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="69e08-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="69e08-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="69e08-120">JavaScript</span></span><br/><span data-ttu-id="69e08-121">или</span><span class="sxs-lookup"><span data-stu-id="69e08-121">or</span></span><br/><span data-ttu-id="69e08-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="69e08-122">TypeScript</span></span> | <span data-ttu-id="69e08-123">Содержит код, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="69e08-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="69e08-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="69e08-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="69e08-125">HTML</span><span class="sxs-lookup"><span data-stu-id="69e08-125">HTML</span></span> | <span data-ttu-id="69e08-126">Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="69e08-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="69e08-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="69e08-127">**./manifest.xml**</span></span> | <span data-ttu-id="69e08-128">XML</span><span class="sxs-lookup"><span data-stu-id="69e08-128">XML</span></span> | <span data-ttu-id="69e08-129">Указывает расположение нескольких файлов, которые используются пользовательскими функциями, например JavaScript, JSON и HTML-файлов.</span><span class="sxs-lookup"><span data-stu-id="69e08-129">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="69e08-130">А также среду выполнения, которую должны использовать пользовательские функции, расположение файлов области задач и командных файлов.</span><span class="sxs-lookup"><span data-stu-id="69e08-130">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="69e08-131">Файл скрипта</span><span class="sxs-lookup"><span data-stu-id="69e08-131">Script file</span></span>

<span data-ttu-id="69e08-132">Файл скрипта (**./src/functions/functions.js** или **./src/functions/functions.ts**) содержит код, определяющий пользовательские функции, и комментарии, определяющие функцию.</span><span class="sxs-lookup"><span data-stu-id="69e08-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="69e08-133">Приведенный ниже код определяет пользовательскую функцию `add`.</span><span class="sxs-lookup"><span data-stu-id="69e08-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="69e08-134">Примечания кода используются для создания файла метаданных JSON с описанием пользовательской функции для Excel.</span><span class="sxs-lookup"><span data-stu-id="69e08-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="69e08-135">Обязательный комментарий `@customfunction` объявлен первым, чтобы указать, что это пользовательская функция.</span><span class="sxs-lookup"><span data-stu-id="69e08-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="69e08-136">Затем объявляются еще два параметра: `first` и `second`, за которыми следуют их свойства `description`.</span><span class="sxs-lookup"><span data-stu-id="69e08-136">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="69e08-137">Наконец, дается описание `returns`.</span><span class="sxs-lookup"><span data-stu-id="69e08-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="69e08-138">Дополнительные сведения о том, какие комментарии являются обязательными для вашей пользовательской функции, см. в статье [Автоматическое создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="69e08-138">For more information about what comments are required for your custom function, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

### <a name="manifest-file"></a><span data-ttu-id="69e08-139">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="69e08-139">Manifest file</span></span>

<span data-ttu-id="69e08-140">Файл манифеста XML для надстройки, определяющий пользовательские функции (**./manifest.xml** в проекте, созданном генератором Yo Office) выполняет следующее:</span><span class="sxs-lookup"><span data-stu-id="69e08-140">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="69e08-141">Определяет пространство имен для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="69e08-141">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="69e08-142">Пространство имен добавляется к пользовательским функциям, чтобы клиенты могли определить ваши функции в рамках надстройки.</span><span class="sxs-lookup"><span data-stu-id="69e08-142">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="69e08-143">Использует уникальные для манифеста пользовательских функций элементы `<ExtensionPoint>` и `<Resources>`.</span><span class="sxs-lookup"><span data-stu-id="69e08-143">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="69e08-144">Эти элементы содержат сведения о расположении JavaScript, JSON и HTML-файлов.</span><span class="sxs-lookup"><span data-stu-id="69e08-144">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="69e08-145">Указывает, какую среду выполнения использовать для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="69e08-145">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="69e08-146">Рекомендуется всегда использовать общую среду выполнения, если нет особой потребности в использовании другой среды, так как общая позволяет делиться данными между функциями и областью задач.</span><span class="sxs-lookup"><span data-stu-id="69e08-146">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span> <span data-ttu-id="69e08-147">Обратите внимание, что использование общей среды выполнения означает, что ваша надстройка будет использовать Internet Explorer 11, а не Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="69e08-147">Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.</span></span>

<span data-ttu-id="69e08-148">Если для создания файлов используется генератор Yo Office, рекомендуется настроить манифест для использования общей среды выполнения, так как это не настроено по умолчанию для этих файлов.</span><span class="sxs-lookup"><span data-stu-id="69e08-148">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="69e08-149">Чтобы изменить манифест, следуйте инструкциям в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="69e08-149">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="69e08-150">Чтобы просмотреть полный рабочий манифест из примера надстройки, см. [этот репозиторий GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="69e08-150">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="69e08-151">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="69e08-151">Coauthoring</span></span>

<span data-ttu-id="69e08-152">Excel для Интернета и Windows, подключенный к подписке Microsoft 365, позволяет использовать совместное редактирование в Excel.</span><span class="sxs-lookup"><span data-stu-id="69e08-152">Excel on the web and on Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="69e08-153">Если в книге используется пользовательская функция, вашему коллеге по совместному редактированию будет предложено загрузить надстройку пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="69e08-153">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="69e08-154">Когда вы оба загрузите надстройку, пользовательская функция поделится результатами с помощью совместного редактирования.</span><span class="sxs-lookup"><span data-stu-id="69e08-154">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="69e08-155">Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="69e08-155">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="next-steps"></a><span data-ttu-id="69e08-156">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="69e08-156">Next steps</span></span>

<span data-ttu-id="69e08-157">Хотите попробовать пользовательские функции?</span><span class="sxs-lookup"><span data-stu-id="69e08-157">Want to try out custom functions?</span></span> <span data-ttu-id="69e08-158">Ознакомьтесь с простым [кратким руководством по началу работы с пользовательскими функциями](../quickstarts/excel-custom-functions-quickstart.md) или с более глубоким [руководством по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md), если вы этого еще не сделали.</span><span class="sxs-lookup"><span data-stu-id="69e08-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="69e08-159">Еще одно простое средство ознакомления с пользовательскими функциями — [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), надстройка, в которой можно экспериментировать с пользовательскими функциями прямо в Excel.</span><span class="sxs-lookup"><span data-stu-id="69e08-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="69e08-160">Вы можете попробовать создать собственные пользовательские функции или поиграть с готовыми примерами.</span><span class="sxs-lookup"><span data-stu-id="69e08-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="69e08-161">См. также</span><span class="sxs-lookup"><span data-stu-id="69e08-161">See also</span></span> 
* [<span data-ttu-id="69e08-162">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="69e08-162">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
* [<span data-ttu-id="69e08-163">Наборы обязательных элементов пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="69e08-163">Custom functions requirement sets</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="69e08-164">Правила именования пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="69e08-164">Custom functions naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="69e08-165">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="69e08-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="69e08-166">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="69e08-166">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
