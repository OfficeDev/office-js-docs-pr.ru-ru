---
ms.date: 12/28/2019
description: Создайте пользовательскую функцию Excel для своей надстройки Office
title: Создание пользовательских функций в Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: f82e99f943b446e546e995f41d104252f8ee6295
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719492"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="592be-103">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="592be-103">Create custom functions in Excel</span></span> 

<span data-ttu-id="592be-104">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="592be-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="592be-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="592be-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="592be-106">В этой статье описано создание специальных функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="592be-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="592be-107">Ниже на анимированном изображении показано, как рабочая книга вызывает функцию, созданную вами с помощью JavaScript или Typescript.</span><span class="sxs-lookup"><span data-stu-id="592be-107">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="592be-108">В этом примере пользовательская функция `=MYFUNCTION.SPHEREVOLUME` рассчитывает объем сферы.</span><span class="sxs-lookup"><span data-stu-id="592be-108">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="592be-109">Приведенный ниже код определяет пользовательскую функцию `=MYFUNCTION.SPHEREVOLUME`.</span><span class="sxs-lookup"><span data-stu-id="592be-109">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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
> <span data-ttu-id="592be-110">В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="592be-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="592be-111">Как определена пользовательская функция в коде</span><span class="sxs-lookup"><span data-stu-id="592be-111">How a custom function is defined in code</span></span>

<span data-ttu-id="592be-112">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, вы обнаружите, что он создает файлы, управляющие вашими функциями, областью задач и надстройкой в целом.</span><span class="sxs-lookup"><span data-stu-id="592be-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="592be-113">Мы сосредоточимся на файлах, которые важны для пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="592be-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="592be-114">Файл</span><span class="sxs-lookup"><span data-stu-id="592be-114">File</span></span> | <span data-ttu-id="592be-115">Формат файла</span><span class="sxs-lookup"><span data-stu-id="592be-115">File format</span></span> | <span data-ttu-id="592be-116">Описание</span><span class="sxs-lookup"><span data-stu-id="592be-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="592be-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="592be-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="592be-118">или</span><span class="sxs-lookup"><span data-stu-id="592be-118">or</span></span><br/><span data-ttu-id="592be-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="592be-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="592be-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="592be-120">JavaScript</span></span><br/><span data-ttu-id="592be-121">или</span><span class="sxs-lookup"><span data-stu-id="592be-121">or</span></span><br/><span data-ttu-id="592be-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="592be-122">TypeScript</span></span> | <span data-ttu-id="592be-123">Содержит код, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="592be-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="592be-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="592be-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="592be-125">HTML</span><span class="sxs-lookup"><span data-stu-id="592be-125">HTML</span></span> | <span data-ttu-id="592be-126">Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="592be-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="592be-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="592be-127">**./manifest.xml**</span></span> | <span data-ttu-id="592be-128">XML</span><span class="sxs-lookup"><span data-stu-id="592be-128">XML</span></span> | <span data-ttu-id="592be-129">Определяет пространство имен для всех пользовательских функций в надстройке и расположение JavaScript и HTML-файлов, которые указаны ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="592be-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="592be-130">Он также перечисляет расположения других файлов, которые могут использоваться надстройкой, например файлы области задач и командные файлы.</span><span class="sxs-lookup"><span data-stu-id="592be-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="592be-131">Файл скрипта</span><span class="sxs-lookup"><span data-stu-id="592be-131">Script file</span></span>

<span data-ttu-id="592be-132">Файл скрипта (**./src/functions/functions.js** или **./src/functions/functions.ts**) содержит код, определяющий пользовательские функции, и комментарии, определяющие функцию.</span><span class="sxs-lookup"><span data-stu-id="592be-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="592be-133">Приведенный ниже код определяет пользовательскую функцию `add`.</span><span class="sxs-lookup"><span data-stu-id="592be-133">The following code defines the custom function `add`.</span></span> <span data-ttu-id="592be-134">Примечания кода используются для создания файла метаданных JSON с описанием пользовательской функции для Excel.</span><span class="sxs-lookup"><span data-stu-id="592be-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="592be-135">Обязательный комментарий `@customfunction` объявлен первым, чтобы указать, что это пользовательская функция.</span><span class="sxs-lookup"><span data-stu-id="592be-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="592be-136">Вы также увидите два объявленных параметра (`first` и `second`), за которыми следуют их свойства `description`.</span><span class="sxs-lookup"><span data-stu-id="592be-136">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="592be-137">Наконец, дается описание `returns`.</span><span class="sxs-lookup"><span data-stu-id="592be-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="592be-138">Дополнительные сведения о том, какие комментарии являются обязательными для вашей пользовательской функции, см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="592be-138">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

<span data-ttu-id="592be-139">Обратите внимание, что файл **functions.html**, который регулирует загрузку среды выполнения пользовательских функций, нужно связать с текущим CDN для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="592be-139">Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions.</span></span> <span data-ttu-id="592be-140">Проекты, подготовленные с текущей версией генератора Yo Office, ссылаются на правильный CDN.</span><span class="sxs-lookup"><span data-stu-id="592be-140">Projects prepared with the current version of the Yo Office generator reference the correct CDN.</span></span> <span data-ttu-id="592be-141">При модернизации предыдущего проекта пользовательской функции от марта 2019 года или более раннего нужно скопировать код, приведенный ниже, на страницу **functions.html**.</span><span class="sxs-lookup"><span data-stu-id="592be-141">If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a><span data-ttu-id="592be-142">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="592be-142">Manifest file</span></span>

<span data-ttu-id="592be-143">XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, который создает генератор Yo Office) и определяет пространство имен для всех пользовательских функций в надстройке, а также расположение файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="592be-143">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span>

<span data-ttu-id="592be-144">Базовая XML-разметка ниже представляет пример элементов `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы активировать пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="592be-144">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="592be-145">Если вы используете генератор Yo Office, созданные файлы пользовательской функции будут содержать более сложный файл манифеста, который можно сравнить в этом [репозитории Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="592be-145">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="592be-146">URL-адреса, указанные в файле манифеста для пользовательских функций файлов JavaScript, JSON и HTML, должны быть общедоступными и иметь один поддомен.</span><span class="sxs-lookup"><span data-stu-id="592be-146">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="592be-147">Функции в Excel имеют в начале пространство имен, указанное в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="592be-147">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="592be-148">Пространство имен функции предшествует названию функции, и они будут разделены точкой.</span><span class="sxs-lookup"><span data-stu-id="592be-148">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="592be-149">Например, чтобы вызвать функцию `ADD42` в ячейке на листе Excel, введите `=CONTOSO.ADD42`, так как `CONTOSO` является пространством имен, а `ADD42` — это имя функции, определяемой в JSON-файл.</span><span class="sxs-lookup"><span data-stu-id="592be-149">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="592be-150">Пространство имен служит в качестве идентификатора для вашей компании или надстройки.</span><span class="sxs-lookup"><span data-stu-id="592be-150">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="592be-151">Пространство имен может содержать только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="592be-151">A namespace can only contain alphanumeric characters and periods.</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="592be-152">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="592be-152">Coauthoring</span></span>

<span data-ttu-id="592be-153">Интернет-версия Excel и Excel для Windows с подпиской на Office 365 позволяют совместно редактировать документы. Эта функция работает с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="592be-153">Excel on the web and Windows connected to an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="592be-154">Если в книге используется пользовательская функция, вашему коллеге будет предложено загрузить надстройку пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="592be-154">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="592be-155">Когда вы оба загрузите надстройку, пользовательская функция поделится результатами с помощью совместного редактирования.</span><span class="sxs-lookup"><span data-stu-id="592be-155">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="592be-156">Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="592be-156">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="592be-157">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="592be-157">Known issues</span></span>

<span data-ttu-id="592be-158">С известными проблемами можно ознакомиться в нашем [репозитории GitHub, посвященном пользовательским функциям в Excel](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="592be-158">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="592be-159">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="592be-159">Next steps</span></span>

<span data-ttu-id="592be-160">Хотите попробовать пользовательские функции?</span><span class="sxs-lookup"><span data-stu-id="592be-160">Want to try out custom functions?</span></span> <span data-ttu-id="592be-161">Ознакомьтесь с простым [кратким руководством по началу работы с пользовательскими функциями](../quickstarts/excel-custom-functions-quickstart.md) или с более глубоким [руководством по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md), если вы этого еще не сделали.</span><span class="sxs-lookup"><span data-stu-id="592be-161">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="592be-162">Еще одно простое средство ознакомления с пользовательскими функциями — [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), надстройка, в которой можно экспериментировать с пользовательскими функциями прямо в Excel.</span><span class="sxs-lookup"><span data-stu-id="592be-162">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="592be-163">Вы можете попробовать создать собственные пользовательские функции или поиграть с готовыми примерами.</span><span class="sxs-lookup"><span data-stu-id="592be-163">You can try out creating your own custom function or play with the provided samples.</span></span>

<span data-ttu-id="592be-164">Готовы узнать больше о возможностях пользовательских функций?</span><span class="sxs-lookup"><span data-stu-id="592be-164">Ready to read more about the capabilities custom functions?</span></span> <span data-ttu-id="592be-165">Ознакомьтесь с обзором [архитектуры пользовательских функций](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="592be-165">Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="592be-166">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="592be-166">See also</span></span> 
* [<span data-ttu-id="592be-167">Требования к настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="592be-167">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="592be-168">Рекомендации по именованию</span><span class="sxs-lookup"><span data-stu-id="592be-168">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="592be-169">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="592be-169">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
