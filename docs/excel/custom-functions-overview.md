---
ms.date: 05/17/2020
description: Создайте пользовательскую функцию Excel для своей надстройки Office
title: Создание пользовательских функций в Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 42ace6208abbd95d0f538345a1f5b5cc15ba1823
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093464"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="f0b9f-103">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="f0b9f-103">Create custom functions in Excel</span></span>

<span data-ttu-id="f0b9f-104">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="f0b9f-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f0b9f-106">Ниже на анимированном изображении показано, как рабочая книга вызывает функцию, созданную вами с помощью JavaScript или Typescript.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="f0b9f-107">В этом примере пользовательская функция `=MYFUNCTION.SPHEREVOLUME` рассчитывает объем сферы.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="f0b9f-108">Приведенный ниже код определяет пользовательскую функцию `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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

> [!NOTE]
> <span data-ttu-id="f0b9f-109">В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-109">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="f0b9f-110">Как определена пользовательская функция в коде</span><span class="sxs-lookup"><span data-stu-id="f0b9f-110">How a custom function is defined in code</span></span>

<span data-ttu-id="f0b9f-111">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания проекта надстройки пользовательских функций Excel, он создает файлы, которые контролируют функции и область задач.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-111">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="f0b9f-112">Мы сосредоточимся на файлах, которые важны для пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="f0b9f-112">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="f0b9f-113">Файл</span><span class="sxs-lookup"><span data-stu-id="f0b9f-113">File</span></span> | <span data-ttu-id="f0b9f-114">Формат файла</span><span class="sxs-lookup"><span data-stu-id="f0b9f-114">File format</span></span> | <span data-ttu-id="f0b9f-115">Описание</span><span class="sxs-lookup"><span data-stu-id="f0b9f-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="f0b9f-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="f0b9f-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="f0b9f-117">или</span><span class="sxs-lookup"><span data-stu-id="f0b9f-117">or</span></span><br/><span data-ttu-id="f0b9f-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="f0b9f-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="f0b9f-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f0b9f-119">JavaScript</span></span><br/><span data-ttu-id="f0b9f-120">или</span><span class="sxs-lookup"><span data-stu-id="f0b9f-120">or</span></span><br/><span data-ttu-id="f0b9f-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f0b9f-121">TypeScript</span></span> | <span data-ttu-id="f0b9f-122">Содержит код, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="f0b9f-123">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="f0b9f-123">**./src/functions/functions.html**</span></span> | <span data-ttu-id="f0b9f-124">HTML</span><span class="sxs-lookup"><span data-stu-id="f0b9f-124">HTML</span></span> | <span data-ttu-id="f0b9f-125">Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-125">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="f0b9f-126">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="f0b9f-126">**./manifest.xml**</span></span> | <span data-ttu-id="f0b9f-127">XML</span><span class="sxs-lookup"><span data-stu-id="f0b9f-127">XML</span></span> | <span data-ttu-id="f0b9f-128">Задает расположение нескольких файлов, используемых настраиваемыми функциями, таких как пользовательские функции, файлы JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-128">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="f0b9f-129">В нем также указаны расположения файлов области задач, командные файлы и указывается, какая среда выполнения должна использоваться вашими пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-129">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="f0b9f-130">Файл скрипта</span><span class="sxs-lookup"><span data-stu-id="f0b9f-130">Script file</span></span>

<span data-ttu-id="f0b9f-131">Файл скрипта (**./src/functions/functions.js** или **./src/functions/functions.ts**) содержит код, определяющий пользовательские функции, и комментарии, определяющие функцию.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-131">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="f0b9f-132">Приведенный ниже код определяет пользовательскую функцию `add`.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-132">The following code defines the custom function `add`.</span></span> <span data-ttu-id="f0b9f-133">Примечания кода используются для создания файла метаданных JSON с описанием пользовательской функции для Excel.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-133">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="f0b9f-134">Обязательный комментарий `@customfunction` объявлен первым, чтобы указать, что это пользовательская функция.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-134">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="f0b9f-135">Затем объявляются два параметра, `first` а `second` затем их `description` Свойства.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-135">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="f0b9f-136">Наконец, дается описание `returns`.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-136">Finally, a `returns` description is given.</span></span> <span data-ttu-id="f0b9f-137">Дополнительные сведения о том, какие комментарии являются обязательными для вашей пользовательской функции, см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="f0b9f-137">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

### <a name="manifest-file"></a><span data-ttu-id="f0b9f-138">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="f0b9f-138">Manifest file</span></span>

<span data-ttu-id="f0b9f-139">XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, созданном генератором Yo Office), выполняет несколько задач:</span><span class="sxs-lookup"><span data-stu-id="f0b9f-139">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="f0b9f-140">Определяет пространство имен для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-140">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="f0b9f-141">Пространство имен добавляется к своим пользовательским функциям, чтобы помочь клиентам определить функции в рамках надстройки.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-141">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="f0b9f-142">Использование `<ExtensionPoint>` и `<Resources>` элементы, которые являются уникальными для манифеста пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-142">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="f0b9f-143">Эти элементы содержат сведения о расположении файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-143">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="f0b9f-144">Указывает, какую среду выполнения использовать для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-144">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="f0b9f-145">Рекомендуется всегда использовать общую среду выполнения, если у вас нет особой необходимости в другой среде выполнения, так как общая среда выполнения позволяет совместно использовать данные между функциями и областью задач.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-145">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span>

<span data-ttu-id="f0b9f-146">Если для создания файлов используется генератор Yo Office, рекомендуется настроить манифест для использования общей среды выполнения, так как это значение по умолчанию не используется для этих файлов.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-146">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="f0b9f-147">Чтобы изменить манифест, следуйте инструкциям в статье [Настройка надстройки Excel, чтобы использовать общую среду выполнения JavaScript](./configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="f0b9f-147">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](./configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="f0b9f-148">Чтобы просмотреть полный рабочий манифест примера надстройки, обратитесь к [репозиторию GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="f0b9f-148">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="f0b9f-149">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="f0b9f-149">Coauthoring</span></span>

<span data-ttu-id="f0b9f-150">Excel в Интернете и Windows, подключенные к подписке Microsoft 365, позволяют совместно редактировать в Excel.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-150">Excel on the web and Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="f0b9f-151">Если ваша книга использует настраиваемую функцию, коллеге соавтору предлагается загрузить надстройку пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-151">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="f0b9f-152">После того как вы загрузили надстройку, настраиваемая функция использует общий доступ к результатам через совместное редактирование.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-152">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="f0b9f-153">Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="f0b9f-153">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="f0b9f-154">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="f0b9f-154">Known issues</span></span>

<span data-ttu-id="f0b9f-155">С известными проблемами можно ознакомиться в нашем [репозитории GitHub, посвященном пользовательским функциям в Excel](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="f0b9f-155">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="f0b9f-156">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="f0b9f-156">Next steps</span></span>

<span data-ttu-id="f0b9f-157">Хотите попробовать пользовательские функции?</span><span class="sxs-lookup"><span data-stu-id="f0b9f-157">Want to try out custom functions?</span></span> <span data-ttu-id="f0b9f-158">Ознакомьтесь с простым [кратким руководством по началу работы с пользовательскими функциями](../quickstarts/excel-custom-functions-quickstart.md) или с более глубоким [руководством по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md), если вы этого еще не сделали.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="f0b9f-159">Еще одно простое средство ознакомления с пользовательскими функциями — [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), надстройка, в которой можно экспериментировать с пользовательскими функциями прямо в Excel.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="f0b9f-160">Вы можете попробовать создать собственные пользовательские функции или поиграть с готовыми примерами.</span><span class="sxs-lookup"><span data-stu-id="f0b9f-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="f0b9f-161">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="f0b9f-161">See also</span></span> 
* [<span data-ttu-id="f0b9f-162">Требования к настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="f0b9f-162">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="f0b9f-163">Рекомендации по именованию</span><span class="sxs-lookup"><span data-stu-id="f0b9f-163">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="f0b9f-164">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="f0b9f-164">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
