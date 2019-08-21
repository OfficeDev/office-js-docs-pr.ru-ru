---
title: Дополнительные концепции программирования с помощью API JavaScript для Excel
description: ''
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: a4639070ed74f9beb757de7c30d1d7e32a3e63fa
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477756"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="ce989-102">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ce989-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="ce989-103">Эта статья является продолжением статьи [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md). В ней описываются более сложные понятия, необходимые для создания сложных надстроек для Excel 2016 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="ce989-103">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="ce989-104">Интерфейсы API Office.js для Excel</span><span class="sxs-lookup"><span data-stu-id="ce989-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="ce989-105">Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="ce989-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="ce989-106">**API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="ce989-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="ce989-107">**Общие API**. Появившиеся в Office 2013 [общие API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="ce989-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="ce989-108">Скорее всего, вы будете разрабатывать большую часть функций надстроек для Excel 2016 или более поздней версии с помощью API JavaScript для Excel, но вам также потребуются объекты из общего API.</span><span class="sxs-lookup"><span data-stu-id="ce989-108">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="ce989-109">Пример:</span><span class="sxs-lookup"><span data-stu-id="ce989-109">For example:</span></span>

- <span data-ttu-id="ce989-p102">[Context](/javascript/api/office/office.context). Объект **Context** представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API. Он состоит из данных конфигурации книги, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`. Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="ce989-p102">[Context](/javascript/api/office/office.context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API. It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`. Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="ce989-113">[Document](/javascript/api/office/office.document). Объект **Document** предоставляет метод `getFileAsync()`, позволяющий скачать файл Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="ce989-113">[Document](/javascript/api/office/office.document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

## <a name="requirement-sets"></a><span data-ttu-id="ce989-114">Наборы требований</span><span class="sxs-lookup"><span data-stu-id="ce989-114">Requirement sets</span></span>

<span data-ttu-id="ce989-p103">Наборы обязательных элементов — это именованные группы элементов API. Надстройка Office может выполнить проверку в среде выполнения или использовать указанные в манифесте наборы обязательных элементов, чтобы определить, поддерживает ли ведущее приложение Office необходимые надстройке API. Сведения о том, какие именно наборы обязательных элементов доступны на каждой поддерживаемой платформе, см. в статье [Наборы обязательных элементов API JavaScript для Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ce989-p103">Requirement sets are named groups of API members. An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs. To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="ce989-118">Проверка поддержки наборов обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="ce989-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="ce989-119">В приведенном ниже примере кода показано, как определить, поддерживает ли ведущее приложение надстройки указанный набор обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="ce989-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="ce989-120">Определение поддержки наборов обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="ce989-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="ce989-p104">С помощью [элемента Requirements](/office/dev/add-ins/reference/manifest/requirements) в манифесте надстройки можно указать минимальные наборы обязательных элементов и/или методы API, необходимые надстройке для активации. Если платформа или ведущее приложение Office не поддерживает наборы обязательных элементов или методы API, указанные в элементе **Requirements** манифеста, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в списке надстроек в разделе **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="ce989-p104">You can use the [Requirements element](/office/dev/add-ins/reference/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate. If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="ce989-123">В приведенном ниже примере кода показан элемент **Requirements** в манифесте надстройки, где указано, что надстройка должна загружаться во всех ведущих приложениях Office, поддерживающих набор обязательных элементов ExcelApi версии 1.3 или выше.</span><span class="sxs-lookup"><span data-stu-id="ce989-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="ce989-124">Чтобы надстройка была доступна на всех платформах ведущего приложения Office, например интернет-версия Excel, Excel для Windows и для iPad, рекомендуем проверять поддержку обязательных элементов в среде выполнения, а не определять поддержку набора обязательных элементов в манифесте.</span><span class="sxs-lookup"><span data-stu-id="ce989-124">To make your add-in available on all platforms of an Office host, such as Excel on Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="ce989-125">Наборы обязательных элементов общего API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="ce989-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="ce989-126">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ce989-126">For information about Common API requirement sets, see [Office Common API requirement sets](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="ce989-127">Загрузка свойств объекта</span><span class="sxs-lookup"><span data-stu-id="ce989-127">Loading the properties of an object</span></span>

<span data-ttu-id="ce989-p105">Вызов метода `load()` для объекта JavaScript в Excel сообщает API, что требуется загрузить объект в память JavaScript при выполнении метода `sync()`. Метод `load()` принимает строку, содержащую разделенные запятыми имена свойств, которые требуется загрузить, или объект, указывающий загружаемые свойства, параметры разбивки на страницы и т. д.</span><span class="sxs-lookup"><span data-stu-id="ce989-p105">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs. The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

> [!NOTE]
> <span data-ttu-id="ce989-p106">Если вызвать метод `load()` для объекта (или коллекции), не указывая параметры, то будут загружены все скалярные свойства объекта (или все скалярные свойства всех объектов в коллекции). Чтобы сократить количество данных, передаваемых между ведущим приложением Excel и надстройкой, следует избегать вызовов метода `load()` без явного указания загружаемых свойств.</span><span class="sxs-lookup"><span data-stu-id="ce989-p106">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded. To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="ce989-132">Сведения о методе</span><span class="sxs-lookup"><span data-stu-id="ce989-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="ce989-133">load(param: объект)</span><span class="sxs-lookup"><span data-stu-id="ce989-133">load(param: object)</span></span>

<span data-ttu-id="ce989-134">Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметрах.</span><span class="sxs-lookup"><span data-stu-id="ce989-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="ce989-135">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="ce989-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="ce989-136">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce989-136">Parameters</span></span>

|<span data-ttu-id="ce989-137">**Параметр**</span><span class="sxs-lookup"><span data-stu-id="ce989-137">**Parameter**</span></span>|<span data-ttu-id="ce989-138">**Тип**</span><span class="sxs-lookup"><span data-stu-id="ce989-138">**Type**</span></span>|<span data-ttu-id="ce989-139">**Описание**</span><span class="sxs-lookup"><span data-stu-id="ce989-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="ce989-140">объект</span><span class="sxs-lookup"><span data-stu-id="ce989-140">object</span></span>|<span data-ttu-id="ce989-p107">Необязательный параметр. Принимает имена свойств в виде строки с разделителями-запятыми или массива. Кроме того, можно передать объект, чтобы задать свойства выделения и навигации (как показано в приведенном ниже примере).</span><span class="sxs-lookup"><span data-stu-id="ce989-p107">Optional. Accepts parameter and relationship names as comma-delimited string or an array. An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="ce989-144">Возвращаемое значение</span><span class="sxs-lookup"><span data-stu-id="ce989-144">Returns</span></span>

<span data-ttu-id="ce989-145">void</span><span class="sxs-lookup"><span data-stu-id="ce989-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="ce989-146">Пример</span><span class="sxs-lookup"><span data-stu-id="ce989-146">Example</span></span>

<span data-ttu-id="ce989-p108">В приведенном ниже примере кода показано, как задать свойства одного диапазона в Excel, скопировав их из другого. Обратите внимание, что для начала необходимо загрузить исходный объект, чтобы можно было получить доступ к значениям его свойств и записать их в целевой диапазон. В этом примере предполагается, что два диапазона (**B2:E2** и **B7:E7**) содержат данные, а их форматирование изначально отличается.</span><span class="sxs-lookup"><span data-stu-id="ce989-p108">The following code sample sets the properties of one Excel range by copying the properties of another range. Note that the source object must be loaded first, before its property values can be accessed and written to the target range. This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a><span data-ttu-id="ce989-150">Загрузка свойств параметров</span><span class="sxs-lookup"><span data-stu-id="ce989-150">Load option properties</span></span>

<span data-ttu-id="ce989-151">Вместо того чтобы передавать строку с разделителями-запятыми или массив при вызове метода `load()`, можно передать объект, содержащий указанные ниже свойства.</span><span class="sxs-lookup"><span data-stu-id="ce989-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="ce989-152">**Свойство**</span><span class="sxs-lookup"><span data-stu-id="ce989-152">**Property**</span></span>|<span data-ttu-id="ce989-153">**Тип**</span><span class="sxs-lookup"><span data-stu-id="ce989-153">**Type**</span></span>|<span data-ttu-id="ce989-154">**Описание**</span><span class="sxs-lookup"><span data-stu-id="ce989-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="ce989-155">объект</span><span class="sxs-lookup"><span data-stu-id="ce989-155">object</span></span>|<span data-ttu-id="ce989-p109">Содержит массив или разделенный запятыми список имен скалярных свойств. Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="ce989-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="ce989-158">объект</span><span class="sxs-lookup"><span data-stu-id="ce989-158">object</span></span>|<span data-ttu-id="ce989-p110">Содержит массив или разделенный запятыми список имен свойств навигации. Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="ce989-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="ce989-161">целое</span><span class="sxs-lookup"><span data-stu-id="ce989-161">int</span></span>| <span data-ttu-id="ce989-p111">Указывает максимальное число элементов в коллекции, которые можно включить в результат. Необязательный параметр. Его можно применять, только если используется параметр нотации объектов.</span><span class="sxs-lookup"><span data-stu-id="ce989-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="ce989-165">int</span><span class="sxs-lookup"><span data-stu-id="ce989-165">int</span></span>|<span data-ttu-id="ce989-p112">Укажите количество элементов в коллекции, которые необходимо пропустить и исключить из результата. Если указан параметр `top`, результирующий набор начнется после пропуска заданного числа элементов. Необязательный. Его можно применять, только если используется параметр нотации объектов.</span><span class="sxs-lookup"><span data-stu-id="ce989-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="ce989-170">В следующем примере кода показано, как загрузить коллекцию листов, выбрав свойства `name` и `address` используемого диапазона для каждого листа в коллекции.</span><span class="sxs-lookup"><span data-stu-id="ce989-170">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="ce989-171">В нем также указано, что следует загружать только пять верхних листов в коллекции.</span><span class="sxs-lookup"><span data-stu-id="ce989-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="ce989-172">Для обработки следующих пяти листов можно указать для атрибутов значения `top: 10` и `skip: 5`.</span><span class="sxs-lookup"><span data-stu-id="ce989-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="ce989-173">Скалярные и навигационные свойства</span><span class="sxs-lookup"><span data-stu-id="ce989-173">Scalar and navigation properties</span></span>

<span data-ttu-id="ce989-174">Существует две категории свойств: **скалярные** и **навигационные**.</span><span class="sxs-lookup"><span data-stu-id="ce989-174">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="ce989-175">К скалярным свойствам относятся назначаемые типы, такие как строки, целые числа и структуры JSON.</span><span class="sxs-lookup"><span data-stu-id="ce989-175">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="ce989-176">Свойства навигации — это объекты и коллекции объектов только для чтения, которым назначаются поля вместо прямого назначения свойства.</span><span class="sxs-lookup"><span data-stu-id="ce989-176">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="ce989-177">Например, элементы `name` и `position` объекта [Worksheet](/javascript/api/excel/excel.worksheet) являются скалярными свойствами, а `protection` и `tables` — свойствами навигации.</span><span class="sxs-lookup"><span data-stu-id="ce989-177">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> <span data-ttu-id="ce989-178">Элемент `prompt` в объекте [DataValidation] является примером скалярного свойства, которое требуется устанавливать с помощью объекта JSON (`dv.prompt = { title: "MyPrompt"}`) вместо настройки подсвойств (`dv.prompt.title = "MyPrompt" // will not set the title`).</span><span class="sxs-lookup"><span data-stu-id="ce989-178">`prompt` on the [DataValidation] object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="ce989-179">Скалярные и навигационные свойства с методом `object.load()`</span><span class="sxs-lookup"><span data-stu-id="ce989-179">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="ce989-p115">При вызове метода `object.load()` без указания параметров загружаются все скалярные свойства объекта. Свойства навигации объекта не загружаются. Кроме того, свойства навигации невозможно загружать напрямую. Вместо этого следует использовать метод `load()`, чтобы ссылаться на отдельные скалярные свойства в нужном свойстве навигации. Например, чтобы загрузить имя шрифта для диапазона, необходимо указать свойства навигации **format** и **font** в качестве пути к свойству **name**:</span><span class="sxs-lookup"><span data-stu-id="ce989-p115">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded. Additionally, navigation properties cannot be loaded directly. Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property. For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="ce989-p116">С помощью API JavaScript для Excel можно задавать скалярные свойства из навигационного свойства по пути к ним. Например, вы можете задать размер шрифта для диапазона с помощью команды `someRange.format.font.size = 10;`. Чтобы задать свойство, необязательно загружать его.</span><span class="sxs-lookup"><span data-stu-id="ce989-p116">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path. For example, you could set the font size for a range by using `someRange.format.font.size = 10;`. You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="ce989-187">Установка свойств объекта</span><span class="sxs-lookup"><span data-stu-id="ce989-187">Setting properties of an object</span></span>

<span data-ttu-id="ce989-p117">Установка свойств объекта с вложенными свойствами навигации может быть трудоемкой задачей. Вместо того чтобы задавать отдельные свойства с помощью путей навигации, как описано выше, вы можете использовать метод `object.set()`, доступный для всех объектов в API JavaScript для Excel. С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.</span><span class="sxs-lookup"><span data-stu-id="ce989-p117">Setting properties on an object with nested navigation properties can be cumbersome. As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API. With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="ce989-p118">Метод `set()` реализован только для объектов API JavaScript для Office в определенных ведущих приложениях, таких как API JavaScript для Excel. Общие API не поддерживают этот метод.</span><span class="sxs-lookup"><span data-stu-id="ce989-p118">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API. The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="ce989-193">set (properties: объект, options: объект)</span><span class="sxs-lookup"><span data-stu-id="ce989-193">set (properties: object, options: object)</span></span>

<span data-ttu-id="ce989-p119">Свойствам объекта, для которого вызывается метод, присваиваются те же значения, что и соответствующим свойствам переданного объекта. Если параметр `properties` является объектом JavaScript, любое свойство в переданном объекте, соответствующее нередактируемому свойству в объекте, для которого вызывается метод, либо игнорируется, либо приводит к возникновению исключения, в зависимости от значения параметра `options`.</span><span class="sxs-lookup"><span data-stu-id="ce989-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="ce989-196">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="ce989-196">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="ce989-197">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce989-197">Parameters</span></span>

|<span data-ttu-id="ce989-198">**Параметр**</span><span class="sxs-lookup"><span data-stu-id="ce989-198">**Parameter**</span></span>|<span data-ttu-id="ce989-199">**Тип**</span><span class="sxs-lookup"><span data-stu-id="ce989-199">**Type**</span></span>|<span data-ttu-id="ce989-200">**Описание**</span><span class="sxs-lookup"><span data-stu-id="ce989-200">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="ce989-201">объект</span><span class="sxs-lookup"><span data-stu-id="ce989-201">object</span></span>|<span data-ttu-id="ce989-202">Либо объект того же типа Office.js, что и объект, для которого вызывается метод, либо объект JavaScript, имена и типы свойств которого повторяют структуру объекта, для которого вызывается метод.</span><span class="sxs-lookup"><span data-stu-id="ce989-202">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="ce989-203">объект</span><span class="sxs-lookup"><span data-stu-id="ce989-203">object</span></span>|<span data-ttu-id="ce989-p120">Необязательный параметр. Может передаваться, только если первый параметр является объектом JavaScript. Объект может содержать следующее свойство: `throwOnReadOnly?: boolean` (по умолчанию — `true`: если переданный объект JavaScript включает нередактируемые свойства, возникает ошибка.)</span><span class="sxs-lookup"><span data-stu-id="ce989-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="ce989-207">Возвращаемое значение</span><span class="sxs-lookup"><span data-stu-id="ce989-207">Returns</span></span>

<span data-ttu-id="ce989-208">void</span><span class="sxs-lookup"><span data-stu-id="ce989-208">void</span></span>

#### <a name="example"></a><span data-ttu-id="ce989-209">Пример</span><span class="sxs-lookup"><span data-stu-id="ce989-209">Example</span></span>

<span data-ttu-id="ce989-p121">В приведенном ниже примере кода показано, как задать несколько свойств формата диапазона, вызвав метод `set()` и передав в него объект JavaScript, имена и типы свойств которого повторяют структуру свойств объекта **Range**. В этом примере предполагается, что данные находятся в диапазоне **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="ce989-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a><span data-ttu-id="ce989-212">Методы &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="ce989-212">&#42;OrNullObject methods</span></span>

<span data-ttu-id="ce989-p122">Многие методы API JavaScript для Excel возвращают исключение, если условие API не соблюдается. Например, если для получения листа указать имя листа, не существующее в книге, то метод `getItem()` вернет исключение `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="ce989-p122">Many Excel JavaScript API methods will return an exception when the condition of the API is not met. For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="ce989-p123">Вместо того чтобы реализовывать сложную логику обработки исключений для такого сценария, можно использовать вариант метода `*OrNullObject`, доступный для нескольких методов в API JavaScript для Excel. Если указанный элемент не существует, метод `*OrNullObject` возвращает нулевой объект (не объект JavaScript `null`), вместо того чтобы возвращать исключение. Например, вы можете вызвать метод `getItemOrNullObject()` для коллекции, например **Worksheets**, чтобы попробовать получить элемент из коллекции. Метод `getItemOrNullObject()` возвращает указанный элемент, если он существует. В противном случае возвращается нулевой объект. Возвращаемый нулевой объект содержит логическое свойство `isNullObject`, с помощью которого можно определить, существует ли объект.</span><span class="sxs-lookup"><span data-stu-id="ce989-p123">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API. An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist. For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection. The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object. The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="ce989-p124">В приведенном ниже примере кода осуществляется попытка получить лист Data с помощью метода `getItemOrNullObject()`. Если метод возвращает нулевой объект, то, прежде чем выполнять какие-либо действия с листом, его необходимо создать.</span><span class="sxs-lookup"><span data-stu-id="ce989-p124">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method. If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a><span data-ttu-id="ce989-222">См. также</span><span class="sxs-lookup"><span data-stu-id="ce989-222">See also</span></span>

* [<span data-ttu-id="ce989-223">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ce989-223">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="ce989-224">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="ce989-224">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="ce989-225">Оптимизация производительности API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ce989-225">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="ce989-226">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ce989-226">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
