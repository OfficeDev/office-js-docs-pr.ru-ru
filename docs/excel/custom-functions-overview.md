---
ms.date: 04/20/2019
description: Создание пользовательских функций в Excel с помощью JavaScript.
title: Создание пользовательских функций в Excel (ознакомительная версия)
localization_priority: Priority
ms.openlocfilehash: 634b76ed90a30c7aa8252da346ba3f95684967a4
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353253"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="983ee-103">Создание пользовательских функций в Excel (ознакомительная версия)</span><span class="sxs-lookup"><span data-stu-id="983ee-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="983ee-104">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel, посредством определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="983ee-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="983ee-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="983ee-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="983ee-106">В этой статье описано создание специальных функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="983ee-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="983ee-107">Ниже продемонстрировано, как конечный пользователь, вставляет настраиваемую функцию в ячейке на листе Excel.</span><span class="sxs-lookup"><span data-stu-id="983ee-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="983ee-108">Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которые пользователь указывает в качестве входных параметров для функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="983ee-109">Приведенный ниже код определяет настраиваемую функцию `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="983ee-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="983ee-110">В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="983ee-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="983ee-111">Компоненты пользовательские функции для надстройки проекта.</span><span class="sxs-lookup"><span data-stu-id="983ee-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="983ee-112">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, вы обнаружите, что он создает файлы, управляющие вашими функциями, областью задач и надстройкой в целом.</span><span class="sxs-lookup"><span data-stu-id="983ee-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="983ee-113">Мы сосредоточимся на файлах, которые важны для пользовательских функций:</span><span class="sxs-lookup"><span data-stu-id="983ee-113">We'll concentrate on the files that are important to custom functions:</span></span> 

| <span data-ttu-id="983ee-114">Файл</span><span class="sxs-lookup"><span data-stu-id="983ee-114">File</span></span> | <span data-ttu-id="983ee-115">Формат файла</span><span class="sxs-lookup"><span data-stu-id="983ee-115">File format</span></span> | <span data-ttu-id="983ee-116">Описание</span><span class="sxs-lookup"><span data-stu-id="983ee-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="983ee-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="983ee-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="983ee-118">или</span><span class="sxs-lookup"><span data-stu-id="983ee-118">or</span></span><br/><span data-ttu-id="983ee-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="983ee-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="983ee-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="983ee-120">JavaScript</span></span><br/><span data-ttu-id="983ee-121">или</span><span class="sxs-lookup"><span data-stu-id="983ee-121">or</span></span><br/><span data-ttu-id="983ee-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="983ee-122">TypeScript</span></span> | <span data-ttu-id="983ee-123">Содержит код, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="983ee-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="983ee-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="983ee-125">HTML</span><span class="sxs-lookup"><span data-stu-id="983ee-125">HTML</span></span> | <span data-ttu-id="983ee-126">Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="983ee-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="983ee-127">**./manifest.xml**</span></span> | <span data-ttu-id="983ee-128">XML</span><span class="sxs-lookup"><span data-stu-id="983ee-128">XML</span></span> | <span data-ttu-id="983ee-129">Определяет пространство имен для всех пользовательских функций в надстройке и расположение JavaScript и HTML-файлов, которые указаны ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="983ee-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="983ee-130">Он также перечисляет расположения других файлов, которые могут использоваться надстройкой, например файлы области задач и командные файлы.</span><span class="sxs-lookup"><span data-stu-id="983ee-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="983ee-131">Файл скрипта</span><span class="sxs-lookup"><span data-stu-id="983ee-131">Script file</span></span>

<span data-ttu-id="983ee-132">Файл скрипта (**./src/functions/functions.js** или **./src/functions/functions.ts** в проекте, созданном генератором Yo Office) содержит код, определяющий пользовательские функции, комментарии, определяющие функцию, и сопоставляет имена пользовательских функций с объектами в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="983ee-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="983ee-133">Указанный ниже код определяет пользовательскую функцию `add` и указывает информацию о сопоставлении для функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-133">The following code defines the custom function `add`  and then specifies association information for the function.</span></span> <span data-ttu-id="983ee-134">Дополнительные сведения о сопоставлении функций см. в статье [Рекомендации по пользовательским функциям](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="983ee-134">For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

<span data-ttu-id="983ee-135">В следующем коде также представлены комментарии кода, определяющие функцию.</span><span class="sxs-lookup"><span data-stu-id="983ee-135">The following code also provides code comments which define the function.</span></span> <span data-ttu-id="983ee-136">Обязательный комментарий `@customfunction` объявлен первым, чтобы указать, что это пользовательская функция.</span><span class="sxs-lookup"><span data-stu-id="983ee-136">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="983ee-137">Вы также увидите два объявленных параметра (`first` и `second`), за которыми следуют их свойства `description`.</span><span class="sxs-lookup"><span data-stu-id="983ee-137">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="983ee-138">Наконец, дается описание `returns`.</span><span class="sxs-lookup"><span data-stu-id="983ee-138">Finally, a `returns` description is given.</span></span> <span data-ttu-id="983ee-139">Дополнительные сведения о том, какие комментарии являются обязательными для вашей пользовательской функции, см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="983ee-139">For more information about what comments are required for your custom function, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a><span data-ttu-id="983ee-140">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="983ee-140">Manifest file</span></span>

<span data-ttu-id="983ee-141">XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, который создает генератор Yo Office) и определяет пространство имен для всех пользовательских функций в надстройке, а также расположение файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="983ee-141">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="983ee-142">Базовая XML-разметка ниже представляет пример элементов `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы активировать пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-142">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="983ee-143">Если вы используете генератор Yo Office, созданные файлы пользовательской функции будут содержать более сложный файл манифеста, который можно сравнить в этом [репозитории Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="983ee-143">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="983ee-144">URL-адреса, указанные в файле манифеста для пользовательских функций файлов JavaScript, JSON и HTML, должны быть общедоступными и иметь один поддомен.</span><span class="sxs-lookup"><span data-stu-id="983ee-144">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
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
> <span data-ttu-id="983ee-145">Функции в Excel имеют в начале пространство имен, указанное в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="983ee-145">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="983ee-146">Пространство имен функции предшествует названию функции, и они будут разделены точкой.</span><span class="sxs-lookup"><span data-stu-id="983ee-146">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="983ee-147">Например, чтобы вызвать функцию `ADD42` в ячейке на листе Excel, введите `=CONTOSO.ADD42`, так как `CONTOSO` является пространством имен, а `ADD42` — это имя функции, определяемой в JSON-файл.</span><span class="sxs-lookup"><span data-stu-id="983ee-147">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="983ee-148">Пространство имен служит в качестве идентификатора для вашей компании или надстройки.</span><span class="sxs-lookup"><span data-stu-id="983ee-148">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="983ee-149">Пространство имен может содержать только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="983ee-149">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="983ee-150">Объявление переменной функции</span><span class="sxs-lookup"><span data-stu-id="983ee-150">Declaring a volatile function</span></span>

<span data-ttu-id="983ee-151">[Переменные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) — это функции, значение которых периодически изменяется, даже если никакой из аргументов функции не меняется.</span><span class="sxs-lookup"><span data-stu-id="983ee-151">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="983ee-152">Эти функции пересчитываются при каждом пересчете в Excel.</span><span class="sxs-lookup"><span data-stu-id="983ee-152">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="983ee-153">К примеру, представьте себе ячейку, вызывающую функцию `NOW`.</span><span class="sxs-lookup"><span data-stu-id="983ee-153">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="983ee-154">При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.</span><span class="sxs-lookup"><span data-stu-id="983ee-154">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="983ee-155">В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="983ee-155">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="983ee-156">Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="983ee-156">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="983ee-157">Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть полезны при обработке дат, времени, случайных чисел и моделировании.</span><span class="sxs-lookup"><span data-stu-id="983ee-157">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="983ee-158">Например, при моделированиях методом Монте-Карло требуется создание случайных входных данных, чтобы определить оптимальное решение.</span><span class="sxs-lookup"><span data-stu-id="983ee-158">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="983ee-159">Чтобы объявить функцию переменной, добавьте `"volatile": true` в объект `options` для функции в файле метаданных JSON, как показано в приведенном ниже примере кода.</span><span class="sxs-lookup"><span data-stu-id="983ee-159">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="983ee-160">Обратите внимание, что функция не может одновременно иметь значения `"streaming": true` и `"volatile": true`. Если оба параметра помечены как `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="983ee-160">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="983ee-161">Состояние сохранения и совместного использования</span><span class="sxs-lookup"><span data-stu-id="983ee-161">Saving and sharing state</span></span>

<span data-ttu-id="983ee-162">Пользовательские функции могут сохранять данные в глобальных переменных JavaScript, которые можно использовать в последующих вызовах.</span><span class="sxs-lookup"><span data-stu-id="983ee-162">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="983ee-163">Сохраненное состояние полезно, когда пользователи вызывают одни и те же настраиваемые функций из более чем одной ячейки, так как все экземпляры функции могут получить доступ к состоянию.</span><span class="sxs-lookup"><span data-stu-id="983ee-163">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="983ee-164">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.</span><span class="sxs-lookup"><span data-stu-id="983ee-164">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="983ee-165">В приведенном ниже примере кода показана реализация вышеописанной функции передачи температуры, сохраняющей состояние с помощью глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="983ee-165">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="983ee-166">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="983ee-166">Note the following about this code:</span></span>

- <span data-ttu-id="983ee-167">Функция `streamTemperature` обновляет значение температуры, которое отображается в ячейке, каждую секунду и использует переменную `savedTemperatures` как источник данных.</span><span class="sxs-lookup"><span data-stu-id="983ee-167">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="983ee-168">Так как `streamTemperature` — это функция потоковой передачи, она реализует обработчик отмены, который будет запускаться, если функция была отменена.</span><span class="sxs-lookup"><span data-stu-id="983ee-168">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="983ee-169">Если пользователь вызывает функцию `streamTemperature` из нескольких ячеек в Excel, функция `streamTemperature` считывает данные из той же самой переменной `savedTemperatures` при каждом запуске.</span><span class="sxs-lookup"><span data-stu-id="983ee-169">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="983ee-170">Функция `refreshTemperature` ежесекундно считывает температуру определенного термометра и сохраняет результат в переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="983ee-170">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="983ee-171">Так как функция `refreshTemperature` недоступна для конечных пользователей в Excel, ее не нужно регистрировать в JSON-файле.</span><span class="sxs-lookup"><span data-stu-id="983ee-171">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="coauthoring"></a><span data-ttu-id="983ee-172">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="983ee-172">Coauthoring</span></span>

<span data-ttu-id="983ee-173">Excel Online и Excel для Windows с подпиской на Office 365 позволяют совместно редактировать документы. Эта функция работает с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="983ee-173">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="983ee-174">Если в книге используется пользовательская функция, вашему коллеге будет предложено загрузить надстройку пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-174">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="983ee-175">Когда вы оба загрузите надстройку, пользовательская функция поделится результатами с помощью совместного редактирования.</span><span class="sxs-lookup"><span data-stu-id="983ee-175">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="983ee-176">Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="983ee-176">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="983ee-177">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="983ee-177">Working with ranges of data</span></span>

<span data-ttu-id="983ee-178">Ваша пользовательская функция может принимать широкий диапазон данных в виде входных параметров или возвращать широкий диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="983ee-178">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="983ee-179">В JavaScript диапазон данных будет иметь вид двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="983ee-179">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="983ee-180">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="983ee-180">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="983ee-181">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="983ee-181">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="983ee-182">Обратите внимание, что в метаданных JSON для данной функции вам следует задать для параметра свойство `type` в `matrix`.</span><span class="sxs-lookup"><span data-stu-id="983ee-182">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="983ee-183">Определение того, какая ячейка вызывала пользовательскую функцию</span><span class="sxs-lookup"><span data-stu-id="983ee-183">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="983ee-184">В некоторых случаях вам потребуется получить адрес ячейки, которая вызывала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="983ee-184">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="983ee-185">Это может быть полезно в следующих типах сценариев:</span><span class="sxs-lookup"><span data-stu-id="983ee-185">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="983ee-186">Форматирование диапазонов: Используйте адрес ячейки в качестве ключа для хранения сведений в [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="983ee-186">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="983ee-187">После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="983ee-187">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="983ee-188">Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `AsyncStorage` с помощью `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="983ee-188">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="983ee-189">Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.</span><span class="sxs-lookup"><span data-stu-id="983ee-189">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="983ee-190">Сведения об адресе ячейки предоставляются только в том случае, если параметру `requiresAddress` присвоено значение `true` в файле метаданных JSON функции.</span><span class="sxs-lookup"><span data-stu-id="983ee-190">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="983ee-191">Ниже приведен пример этого при создании JSON-файла вручную.</span><span class="sxs-lookup"><span data-stu-id="983ee-191">The following sample gives an example of this if you were to write this JSON file by hand.</span></span> <span data-ttu-id="983ee-192">Вы также можете использовать тег `@requiresAddress` при автоматическом создании JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="983ee-192">You can also use the `@requiresAddress` tag if automatically generating your JSON file.</span></span> <span data-ttu-id="983ee-193">Дополнительные сведения см. в статье [Автоматическое создание JSON](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="983ee-193">For more details, see [JSON Autogeneration](custom-functions-json-autogeneration.md).</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="983ee-194">Чтобы найти адрес ячейки, в файл скрипта (**./src/functions/functions.js** или **./src/functions/functions.ts**) потребуется также добавить функцию `getAddress`.</span><span class="sxs-lookup"><span data-stu-id="983ee-194">In the script file (**./src/functions/functions.js** or **./src/functions/functions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="983ee-195">В этой функции можно использовать параметры, как показано в примере ниже в виде `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="983ee-195">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="983ee-196">В качестве последнего параметра всегда будет использоваться `invocationContext` — объект, содержащий расположение ячейки, которое передает приложение Excel, если параметру `requiresAddress` присвоено значение `true` в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="983ee-196">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="983ee-197">По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="983ee-197">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="983ee-198">Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="983ee-198">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="983ee-199">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="983ee-199">Known issues</span></span>

<span data-ttu-id="983ee-200">С известными проблемами можно ознакомиться в нашем [репозитории GitHub, посвященном пользовательским функциям в Excel](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="983ee-200">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="983ee-201">См. также</span><span class="sxs-lookup"><span data-stu-id="983ee-201">See also</span></span>

* [<span data-ttu-id="983ee-202">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="983ee-202">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="983ee-203">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="983ee-203">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="983ee-204">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="983ee-204">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="983ee-205">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="983ee-205">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="983ee-206">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="983ee-206">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="983ee-207">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="983ee-207">Custom functions debugging</span></span>](custom-functions-debugging.md)
