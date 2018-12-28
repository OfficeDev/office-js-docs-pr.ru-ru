---
title: Дополнительные концепции программирования с помощью API JavaScript для Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 7abc6b692ed6d72924e7ebda47a8198fd85a4aa0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457882"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="2e016-102">Дополнительные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2e016-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="2e016-103">Эта статья является продолжением статьи [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md). В ней описываются более сложные понятия, необходимые для создания сложных надстроек для Excel 2016 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="2e016-103">This article builds upon the information in [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="2e016-104">Интерфейсы API Office.js для Excel</span><span class="sxs-lookup"><span data-stu-id="2e016-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="2e016-105">Надстройка Excel взаимодействует с объектами в Excel с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="2e016-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="2e016-106">**API JavaScript для Excel**. Появившийся в Office 2016 [API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к листам, диапазонам, таблицам, диаграммам и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="2e016-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="2e016-107">**Общие API**. Появившиеся в Office 2013 [общие API](../reference/javascript-api-for-office.md) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов ведущих приложений, например Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2e016-107">**Common APIs**: Introduced with Office 2013, the common APIs (also referred to as the [Shared API](../reference/javascript-api-for-office.md)) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="2e016-108">Скорее всего, вы будете разрабатывать большую часть функций надстроек для Excel 2016 или более поздней версии с помощью API JavaScript для Excel, но вам также потребуются объекты из общего API.</span><span class="sxs-lookup"><span data-stu-id="2e016-108">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016, you'll also use objects in the Shared API.</span></span> <span data-ttu-id="2e016-109">Например:</span><span class="sxs-lookup"><span data-stu-id="2e016-109">For example:</span></span>

- <span data-ttu-id="2e016-110">[Context](https://docs.microsoft.com/javascript/api/office/office.context). Объект **Context** представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API.</span><span class="sxs-lookup"><span data-stu-id="2e016-110">[Context](https://docs.microsoft.com/javascript/api/office/office.context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="2e016-111">Он состоит из данных конфигурации книги, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`.</span><span class="sxs-lookup"><span data-stu-id="2e016-111">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="2e016-112">Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="2e016-112">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="2e016-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document). Объект **Document** предоставляет метод `getFileAsync()`, позволяющий скачать файл Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="2e016-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="2e016-114">Наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="2e016-114">Requirement sets</span></span>

<span data-ttu-id="2e016-115">Наборы обязательных элементов — это именованные группы элементов API.</span><span class="sxs-lookup"><span data-stu-id="2e016-115">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="2e016-116">Надстройка Office может выполнить проверку в среде выполнения или использовать указанные в манифесте наборы обязательных элементов, чтобы определить, поддерживает ли ведущее приложение Office необходимые надстройке API.</span><span class="sxs-lookup"><span data-stu-id="2e016-116">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="2e016-117">Сведения о том, какие именно наборы обязательных элементов доступны на каждой поддерживаемой платформе, см. в статье [Наборы обязательных элементов API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2e016-117">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="2e016-118">Проверка поддержки наборов обязательных элементов в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="2e016-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="2e016-119">В приведенном ниже примере кода показано, как определить, поддерживает ли ведущее приложение надстройки указанный набор обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="2e016-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="2e016-120">Определение поддержки наборов обязательных элементов в манифесте</span><span class="sxs-lookup"><span data-stu-id="2e016-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="2e016-121">С помощью [элемента Requirements](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements) в манифесте надстройки можно указать минимальные наборы обязательных элементов и/или методы API, необходимые надстройке для активации.</span><span class="sxs-lookup"><span data-stu-id="2e016-121">You can use the [Requirements element](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="2e016-122">Если платформа или ведущее приложение Office не поддерживает наборы обязательных элементов или методы API, указанные в элементе **Requirements** манифеста, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в списке надстроек в разделе **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="2e016-122">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="2e016-123">В приведенном ниже примере кода показан элемент **Requirements** в манифесте надстройки, где указано, что надстройка должна загружаться во всех ведущих приложениях Office, поддерживающих набор обязательных элементов ExcelApi версии 1.3 или выше.</span><span class="sxs-lookup"><span data-stu-id="2e016-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="2e016-124">Чтобы надстройка была доступна на всех платформах ведущего приложения Office, например Excel для Windows, Excel Online и Excel для iPad, рекомендуем проверять поддержку обязательных элементов в среде выполнения, а не определять поддержку набора обязательных элементов в манифесте.</span><span class="sxs-lookup"><span data-stu-id="2e016-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="2e016-125">Наборы обязательных элементов общего API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="2e016-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="2e016-126">Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2e016-126">For information about common API requirement sets, see [Office common API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="2e016-127">Загрузка свойств объекта</span><span class="sxs-lookup"><span data-stu-id="2e016-127">Loading the properties of an object</span></span>

<span data-ttu-id="2e016-128">Вызов метода `load()` для объекта JavaScript в Excel сообщает API, что требуется загрузить объект в память JavaScript при выполнении метода `sync()`.</span><span class="sxs-lookup"><span data-stu-id="2e016-128">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="2e016-129">Метод `load()` принимает строку, содержащую разделенные запятыми имена свойств, которые требуется загрузить, или объект, указывающий загружаемые свойства, параметры разбивки на страницы и т. д.</span><span class="sxs-lookup"><span data-stu-id="2e016-129">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="2e016-130">Если вызвать метод `load()` для объекта (или коллекции), не указывая параметры, то будут загружены все скалярные свойства объекта (или все скалярные свойства всех объектов в коллекции).</span><span class="sxs-lookup"><span data-stu-id="2e016-130">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="2e016-131">Чтобы сократить количество данных, передаваемых между ведущим приложением Excel и надстройкой, следует избегать вызовов метода `load()` без явного указания загружаемых свойств.</span><span class="sxs-lookup"><span data-stu-id="2e016-131">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="2e016-132">Сведения о методе</span><span class="sxs-lookup"><span data-stu-id="2e016-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="2e016-133">load(param: объект)</span><span class="sxs-lookup"><span data-stu-id="2e016-133">load(param: object)</span></span>

<span data-ttu-id="2e016-134">Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметрах.</span><span class="sxs-lookup"><span data-stu-id="2e016-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="2e016-135">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2e016-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="2e016-136">Параметры</span><span class="sxs-lookup"><span data-stu-id="2e016-136">Parameters</span></span>

|<span data-ttu-id="2e016-137">**Параметр**</span><span class="sxs-lookup"><span data-stu-id="2e016-137">**Parameter**</span></span>|<span data-ttu-id="2e016-138">**Тип**</span><span class="sxs-lookup"><span data-stu-id="2e016-138">**Type**</span></span>|<span data-ttu-id="2e016-139">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2e016-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="2e016-140">объект</span><span class="sxs-lookup"><span data-stu-id="2e016-140">object</span></span>|<span data-ttu-id="2e016-141">Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="2e016-141">Optional.</span></span> <span data-ttu-id="2e016-142">Принимает имена параметров и связей в виде строки с разделителями-запятыми или массива.</span><span class="sxs-lookup"><span data-stu-id="2e016-142">Accepts parameter and relationship names as comma-delimited string or an array.</span></span> <span data-ttu-id="2e016-143">Кроме того, можно передать объект, чтобы задать свойства выделения и навигации (как показано в приведенном ниже примере).</span><span class="sxs-lookup"><span data-stu-id="2e016-143">An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="2e016-144">Возвращаемое значение</span><span class="sxs-lookup"><span data-stu-id="2e016-144">Returns</span></span>

<span data-ttu-id="2e016-145">void</span><span class="sxs-lookup"><span data-stu-id="2e016-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="2e016-146">Пример</span><span class="sxs-lookup"><span data-stu-id="2e016-146">Example</span></span>

<span data-ttu-id="2e016-147">В приведенном ниже примере кода показано, как задать свойства одного диапазона в Excel, скопировав их из другого.</span><span class="sxs-lookup"><span data-stu-id="2e016-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="2e016-148">Обратите внимание, что для начала необходимо загрузить исходный объект, чтобы можно было получить доступ к значениям его свойств и записать их в целевой диапазон.</span><span class="sxs-lookup"><span data-stu-id="2e016-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="2e016-149">В этом примере предполагается, что два диапазона (**B2:E2** и **B7:E7**) содержат данные, а их форматирование изначально отличается.</span><span class="sxs-lookup"><span data-stu-id="2e016-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="2e016-150">Загрузка свойств параметров</span><span class="sxs-lookup"><span data-stu-id="2e016-150">Load option properties</span></span>

<span data-ttu-id="2e016-151">Вместо того чтобы передавать строку с разделителями-запятыми или массив при вызове метода `load()`, можно передать объект, содержащий указанные ниже свойства.</span><span class="sxs-lookup"><span data-stu-id="2e016-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="2e016-152">**Свойство**</span><span class="sxs-lookup"><span data-stu-id="2e016-152">**Property**</span></span>|<span data-ttu-id="2e016-153">**Тип**</span><span class="sxs-lookup"><span data-stu-id="2e016-153">**Type**</span></span>|<span data-ttu-id="2e016-154">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2e016-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="2e016-155">object</span><span class="sxs-lookup"><span data-stu-id="2e016-155">object</span></span>|<span data-ttu-id="2e016-p109">Содержит массив или разделенный запятыми список имен параметров и связей. Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="2e016-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="2e016-158">object</span><span class="sxs-lookup"><span data-stu-id="2e016-158">object</span></span>|<span data-ttu-id="2e016-p110">Содержит массив или разделенный запятыми список имен связей. Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="2e016-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="2e016-161">int</span><span class="sxs-lookup"><span data-stu-id="2e016-161">int</span></span>| <span data-ttu-id="2e016-p111">Указывает максимальное число элементов в коллекции, которые можно включить в результат. Необязательный параметр. Его можно применять, только если используется параметр нотации объектов.</span><span class="sxs-lookup"><span data-stu-id="2e016-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="2e016-165">int</span><span class="sxs-lookup"><span data-stu-id="2e016-165">int</span></span>|<span data-ttu-id="2e016-p112">Укажите количество элементов в коллекции, которые необходимо пропустить и исключить из результата. Если указан параметр `top`, результирующий набор начнется после пропуска заданного числа элементов. Необязательный. Его можно применять, только если используется параметр нотации объектов.</span><span class="sxs-lookup"><span data-stu-id="2e016-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="2e016-170">В приведенном ниже примере кода показано, как загрузить коллекцию листов, выбрав свойства `name` и `address` используемого диапазона для каждого листа в коллекции.</span><span class="sxs-lookup"><span data-stu-id="2e016-170">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="2e016-171">В нем также указано, что следует загружать только пять верхних листов в коллекции.</span><span class="sxs-lookup"><span data-stu-id="2e016-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="2e016-172">Для обработки следующих пяти листов можно указать для атрибутов значения `top: 10` и `skip: 5`.</span><span class="sxs-lookup"><span data-stu-id="2e016-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="2e016-173">Скалярные и навигационные свойства</span><span class="sxs-lookup"><span data-stu-id="2e016-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="2e016-174">Вы могли заметить, что в справочной документации по API JavaScript для Excel элементы объектов делятся на две категории: **свойства** и **связи**.</span><span class="sxs-lookup"><span data-stu-id="2e016-174">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**.</span></span> <span data-ttu-id="2e016-175">Свойство объекта — это скалярный элемент, например строка, целое число или логическое значение, а связь объекта (другое название — свойство навигации) — это элемент, представляющий собой объект или их коллекцию.</span><span class="sxs-lookup"><span data-stu-id="2e016-175">A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects.</span></span> <span data-ttu-id="2e016-176">Например, элементы `name` и `position` объекта [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) являются скалярными свойствами, а `protection` и `tables` — связями (свойствами навигации).</span><span class="sxs-lookup"><span data-stu-id="2e016-176">For example, `name` and `position` members on the [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="2e016-177">Скалярные и навигационные свойства с методом `object.load()`</span><span class="sxs-lookup"><span data-stu-id="2e016-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="2e016-178">При вызове метода `object.load()` без указания параметров загружаются все скалярные свойства объекта. Свойства навигации объекта не загружаются.</span><span class="sxs-lookup"><span data-stu-id="2e016-178">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="2e016-179">Кроме того, свойства навигации невозможно загружать напрямую.</span><span class="sxs-lookup"><span data-stu-id="2e016-179">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="2e016-180">Вместо этого следует использовать метод `load()`, чтобы ссылаться на отдельные скалярные свойства в нужном свойстве навигации.</span><span class="sxs-lookup"><span data-stu-id="2e016-180">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="2e016-181">Например, чтобы загрузить имя шрифта для диапазона, необходимо указать свойства навигации **format** и **font** в качестве пути к свойству **name**:</span><span class="sxs-lookup"><span data-stu-id="2e016-181">For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="2e016-182">С помощью API JavaScript для Excel можно задавать скалярные свойства из навигационного свойства по пути к ним.</span><span class="sxs-lookup"><span data-stu-id="2e016-182">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="2e016-183">Например, вы можете задать размер шрифта для диапазона с помощью команды `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="2e016-183">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="2e016-184">Чтобы задать свойство, необязательно загружать его.</span><span class="sxs-lookup"><span data-stu-id="2e016-184">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="2e016-185">Установка свойств объекта</span><span class="sxs-lookup"><span data-stu-id="2e016-185">Setting properties of an object</span></span>

<span data-ttu-id="2e016-186">Установка свойств объекта с вложенными свойствами навигации может быть трудоемкой задачей.</span><span class="sxs-lookup"><span data-stu-id="2e016-186">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="2e016-187">Вместо того чтобы задавать отдельные свойства с помощью путей навигации, как описано выше, вы можете использовать метод `object.set()`, доступный для всех объектов в API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="2e016-187">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="2e016-188">С помощью этого метода можно задать сразу несколько свойств объекта, передавая другой объект того же типа Office.js или объект JavaScript со свойствами, сходными по структуре со свойствами объекта, для которого вызывается метод.</span><span class="sxs-lookup"><span data-stu-id="2e016-188">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="2e016-189">Метод `set()` реализован только для объектов API JavaScript для Office в определенных ведущих приложениях, таких как API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="2e016-189">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="2e016-190">Общие API не поддерживают этот метод.</span><span class="sxs-lookup"><span data-stu-id="2e016-190">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="2e016-191">set (properties: объект, options: объект)</span><span class="sxs-lookup"><span data-stu-id="2e016-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="2e016-p119">Свойствам объекта, для которого вызывается метод, присваиваются те же значения, что и соответствующим свойствам переданного объекта. Если параметр `properties` является объектом JavaScript, любое свойство в переданном объекте, соответствующее нередактируемому свойству в объекте, для которого вызывается метод, либо игнорируется, либо приводит к возникновению исключения, в зависимости от значения параметра `options`.</span><span class="sxs-lookup"><span data-stu-id="2e016-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="2e016-194">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2e016-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="2e016-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="2e016-195">Parameters</span></span>

|<span data-ttu-id="2e016-196">**Параметр**</span><span class="sxs-lookup"><span data-stu-id="2e016-196">**Parameter**</span></span>|<span data-ttu-id="2e016-197">**Тип**</span><span class="sxs-lookup"><span data-stu-id="2e016-197">**Type**</span></span>|<span data-ttu-id="2e016-198">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2e016-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="2e016-199">объект</span><span class="sxs-lookup"><span data-stu-id="2e016-199">object</span></span>|<span data-ttu-id="2e016-200">Либо объект того же типа Office.js, что и объект, для которого вызывается метод, либо объект JavaScript, имена и типы свойств которого повторяют структуру объекта, для которого вызывается метод.</span><span class="sxs-lookup"><span data-stu-id="2e016-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="2e016-201">объект</span><span class="sxs-lookup"><span data-stu-id="2e016-201">object</span></span>|<span data-ttu-id="2e016-p120">Необязательный параметр. Может передаваться, только если первый параметр является объектом JavaScript. Объект может содержать следующее свойство: `throwOnReadOnly?: boolean` (по умолчанию — `true`: если переданный объект JavaScript включает нередактируемые свойства, возникает ошибка.)</span><span class="sxs-lookup"><span data-stu-id="2e016-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="2e016-205">Возвращаемое значение</span><span class="sxs-lookup"><span data-stu-id="2e016-205">Returns</span></span>

<span data-ttu-id="2e016-206">void</span><span class="sxs-lookup"><span data-stu-id="2e016-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="2e016-207">Пример</span><span class="sxs-lookup"><span data-stu-id="2e016-207">Example</span></span>

<span data-ttu-id="2e016-p121">В приведенном ниже примере кода показано, как задать несколько свойств формата диапазона, вызвав метод `set()` и передав в него объект JavaScript, имена и типы свойств которого повторяют структуру свойств объекта **Range**. В этом примере предполагается, что данные находятся в диапазоне **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="2e016-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="2e016-210">Методы &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="2e016-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="2e016-211">Многие методы API JavaScript для Excel возвращают исключение, если условие API не соблюдается.</span><span class="sxs-lookup"><span data-stu-id="2e016-211">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="2e016-212">Например, если для получения листа указать имя листа, не существующее в книге, то метод `getItem()` вернет исключение `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="2e016-212">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="2e016-213">Вместо того чтобы реализовывать сложную логику обработки исключений для такого сценария, можно использовать вариант метода `*OrNullObject`, доступный для нескольких методов в API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="2e016-213">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="2e016-214">Если указанный элемент не существует, метод `*OrNullObject` возвращает нулевой объект (не объект JavaScript `null`), вместо того чтобы возвращать исключение.</span><span class="sxs-lookup"><span data-stu-id="2e016-214">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="2e016-215">Например, вы можете вызвать метод `getItemOrNullObject()` для коллекции, например **Worksheets**, чтобы попробовать получить элемент из коллекции.</span><span class="sxs-lookup"><span data-stu-id="2e016-215">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="2e016-216">Метод `getItemOrNullObject()` возвращает указанный элемент, если он существует. В противном случае возвращается нулевой объект.</span><span class="sxs-lookup"><span data-stu-id="2e016-216">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="2e016-217">Возвращаемый нулевой объект содержит логическое свойство `isNullObject`, с помощью которого можно определить, существует ли объект.</span><span class="sxs-lookup"><span data-stu-id="2e016-217">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="2e016-218">В приведенном ниже примере кода осуществляется попытка получить лист Data с помощью метода `getItemOrNullObject()`.</span><span class="sxs-lookup"><span data-stu-id="2e016-218">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="2e016-219">Если метод возвращает нулевой объект, то, прежде чем выполнять какие-либо действия с листом, его необходимо создать.</span><span class="sxs-lookup"><span data-stu-id="2e016-219">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2e016-220">См. также</span><span class="sxs-lookup"><span data-stu-id="2e016-220">See also</span></span>
 
* [<span data-ttu-id="2e016-221">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2e016-221">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* <span data-ttu-id="2e016-222">
  [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)</span><span class="sxs-lookup"><span data-stu-id="2e016-222">[Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)</span></span>
* [<span data-ttu-id="2e016-223">Оптимизация производительности API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2e016-223">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="2e016-224">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2e016-224">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
