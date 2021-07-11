---
title: Устранение Excel надстройки
description: Узнайте, как устранить ошибки разработки в Excel надстройки.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: cb622a1805be7bec61168ab37a41709a57075788
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349443"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="289ad-103">Устранение Excel надстройки</span><span class="sxs-lookup"><span data-stu-id="289ad-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="289ad-104">В этой статье обсуждаются проблемы устранения неполадок, которые уникальны для Excel.</span><span class="sxs-lookup"><span data-stu-id="289ad-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="289ad-105">Используйте средство обратной связи в нижней части страницы, чтобы предложить другие проблемы, которые можно добавить в статью.</span><span class="sxs-lookup"><span data-stu-id="289ad-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="289ad-106">Ограничения API при активных переключателях книг</span><span class="sxs-lookup"><span data-stu-id="289ad-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="289ad-107">Надстройки для Excel предназначены для работы с одной книгой одновременно.</span><span class="sxs-lookup"><span data-stu-id="289ad-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="289ad-108">Ошибки могут возникать, когда книга, которая отделена от книги, которая работает надстройка получает фокус.</span><span class="sxs-lookup"><span data-stu-id="289ad-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="289ad-109">Это происходит только в том случае, если конкретные методы находятся в процессе призыва при смене фокуса.</span><span class="sxs-lookup"><span data-stu-id="289ad-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="289ad-110">На следующие API влияет этот переключатель книги.</span><span class="sxs-lookup"><span data-stu-id="289ad-110">The following APIs are affected by this workbook switch.</span></span>

|<span data-ttu-id="289ad-111">API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="289ad-111">Excel JavaScript API</span></span> | <span data-ttu-id="289ad-112">Ошибка, брошенная</span><span class="sxs-lookup"><span data-stu-id="289ad-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="289ad-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="289ad-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="289ad-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="289ad-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="289ad-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="289ad-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="289ad-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="289ad-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="289ad-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="289ad-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="289ad-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="289ad-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="289ad-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="289ad-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="289ad-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="289ad-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="289ad-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="289ad-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="289ad-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="289ad-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="289ad-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="289ad-129">Это касается только нескольких Excel книг, открытых на Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="289ad-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="289ad-130">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="289ad-130">Coauthoring</span></span>

<span data-ttu-id="289ad-131">См. [в](co-authoring-in-excel-add-ins.md) Excel надстройки для шаблонов, которые можно использовать с событиями в среде совместной работы.</span><span class="sxs-lookup"><span data-stu-id="289ad-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="289ad-132">В статье также обсуждаются потенциальные конфликты слияния при использовании определенных API, например [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="289ad-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="289ad-133">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="289ad-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="289ad-134">События привязки возвращают `Binding` временные obects</span><span class="sxs-lookup"><span data-stu-id="289ad-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="289ad-135">Оба [bindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) и [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) возвращают временный объект, содержащий ID объекта, который поднял `Binding` `Binding` событие.</span><span class="sxs-lookup"><span data-stu-id="289ad-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="289ad-136">Используйте этот ID для `BindingCollection.getItem(id)` получения `Binding` объекта, который поднял событие.</span><span class="sxs-lookup"><span data-stu-id="289ad-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="289ad-137">В следующем примере кода показано, как использовать этот временный код привязки для получения связанного `Binding` объекта.</span><span class="sxs-lookup"><span data-stu-id="289ad-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="289ad-138">В примере слушателю событий назначена привязка.</span><span class="sxs-lookup"><span data-stu-id="289ad-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="289ad-139">При запуске события слушатель вызывает `getBindingId` `onDataChanged` метод.</span><span class="sxs-lookup"><span data-stu-id="289ad-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="289ad-140">Метод использует ID временного объекта для получения объекта, `getBindingId` `Binding` который поднял `Binding` событие.</span><span class="sxs-lookup"><span data-stu-id="289ad-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="289ad-141">Формат `useStandardHeight` ячейки `useStandardWidth` и проблемы</span><span class="sxs-lookup"><span data-stu-id="289ad-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="289ad-142">Свойство [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) не работает должным образом `CellPropertiesFormat` в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="289ad-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="289ad-143">Из-за проблемы в пользовательском интерфейсе Excel в Интернете, задав свойство для нечетких расчетов высоты `useStandardHeight` `true` на этой платформе.</span><span class="sxs-lookup"><span data-stu-id="289ad-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="289ad-144">Например, стандартная высота **14** изменена до **14,25** в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="289ad-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="289ad-145">На всех платформах свойства [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) и [UseStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) предназначены только для `CellPropertiesFormat` `true` .</span><span class="sxs-lookup"><span data-stu-id="289ad-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="289ad-146">Настройка этих свойств не `false` влияет.</span><span class="sxs-lookup"><span data-stu-id="289ad-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="289ad-147">Метод `getImage` Range, неподтвердимый Excel для Mac</span><span class="sxs-lookup"><span data-stu-id="289ad-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="289ad-148">Метод [Range getImage](/javascript/api/excel/excel.range#getImage__) в настоящее время не поддерживается в Excel для Mac.</span><span class="sxs-lookup"><span data-stu-id="289ad-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="289ad-149">См. [в #235 OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues/235) Issue для текущего состояния.</span><span class="sxs-lookup"><span data-stu-id="289ad-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="289ad-150">Ограничение возвращаемого символа диапазона</span><span class="sxs-lookup"><span data-stu-id="289ad-150">Range return character limit</span></span>

<span data-ttu-id="289ad-151">Методы [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) и [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) имеют ограничение строки адресов в 8192 символа.</span><span class="sxs-lookup"><span data-stu-id="289ad-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="289ad-152">При превышении этого ограничения строка адресов будет усечена до 8192 символов.</span><span class="sxs-lookup"><span data-stu-id="289ad-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="289ad-153">См. также</span><span class="sxs-lookup"><span data-stu-id="289ad-153">See also</span></span>

- [<span data-ttu-id="289ad-154">Устранение ошибок разработки с Office надстройки</span><span class="sxs-lookup"><span data-stu-id="289ad-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="289ad-155">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="289ad-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
