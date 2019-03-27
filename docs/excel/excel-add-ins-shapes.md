---
title: Работать с фигурами с помощью API JavaScript для Excel
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: e4d01c387fff01d68cb26369240a1e06e723a54c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30926679"
---
# <a name="work-with-shapes-using-the-excel-javascript-api-preview"></a><span data-ttu-id="b030d-102">Работать с фигурами с помощью API JavaScript для Excel (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="b030d-102">Work with Shapes using the Excel JavaScript API (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="b030d-103">API, обсуждаемые в этой статье, в настоящее время доступны только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="b030d-103">The APIs discussed in this article are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="b030d-104">Excel определяет фигуры как объекты, расположенные в графическом слое Excel.</span><span class="sxs-lookup"><span data-stu-id="b030d-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="b030d-105">Это означает, что все за прев ячейке ячейка является фигурой.</span><span class="sxs-lookup"><span data-stu-id="b030d-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="b030d-106">В этой статье описывается, как использовать геометрические фигуры, линии и изображения в сочетании с API [Shape]/жаваскрипт/АПИ/ексцел/ексцел.Шапе) и [ShapeCollection](/javascript/api/excel/excel.shapecollection) .</span><span class="sxs-lookup"><span data-stu-id="b030d-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape]/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="b030d-107">[Диаграммы](/javascript/api/excel/excel.chart) рассматриваются в собственной статье [работать с диаграммами с помощью API JavaScript для Excel]] (Excel-Add-ins-Charts.md)).</span><span class="sxs-lookup"><span data-stu-id="b030d-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with Charts using the Excel JavaScript API]](excel-add-ins-charts.md)).</span></span>

## <a name="create-shapes"></a><span data-ttu-id="b030d-108">Создание фигур</span><span class="sxs-lookup"><span data-stu-id="b030d-108">Create shapes</span></span>

<span data-ttu-id="b030d-109">Фигуры создаются и хранятся в коллекции фигур листа (`Worksheet.shapes`).</span><span class="sxs-lookup"><span data-stu-id="b030d-109">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="b030d-110">`ShapeCollection`в этой `.add*` цели есть несколько способов.</span><span class="sxs-lookup"><span data-stu-id="b030d-110">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="b030d-111">Все фигуры имеют имена и идентификаторы, созданные для них при добавлении в коллекцию.</span><span class="sxs-lookup"><span data-stu-id="b030d-111">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="b030d-112">Это свойства `name` и `id` , соответственно, свойства.</span><span class="sxs-lookup"><span data-stu-id="b030d-112">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="b030d-113">`name`может быть задано надстройкой для упрощения поиска с помощью `ShapeCollection.getItem(name)` метода.</span><span class="sxs-lookup"><span data-stu-id="b030d-113">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="b030d-114">С помощью связанного метода добавляются следующие типы фигур:</span><span class="sxs-lookup"><span data-stu-id="b030d-114">The following types of shapes are added using the associated method:</span></span>

| <span data-ttu-id="b030d-115">Shape</span><span class="sxs-lookup"><span data-stu-id="b030d-115">Shape</span></span> | <span data-ttu-id="b030d-116">Добавление метода</span><span class="sxs-lookup"><span data-stu-id="b030d-116">Add Method</span></span> | <span data-ttu-id="b030d-117">Подпись</span><span class="sxs-lookup"><span data-stu-id="b030d-117">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="b030d-118">ГеоМетрическая фигура</span><span class="sxs-lookup"><span data-stu-id="b030d-118">Geometric Shape</span></span> | [<span data-ttu-id="b030d-119">Адджеометрикшапе</span><span class="sxs-lookup"><span data-stu-id="b030d-119">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="b030d-120">Изображение (JPEG или PNG)</span><span class="sxs-lookup"><span data-stu-id="b030d-120">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="b030d-121">Аддимаже</span><span class="sxs-lookup"><span data-stu-id="b030d-121">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="b030d-122">Line</span><span class="sxs-lookup"><span data-stu-id="b030d-122">Line</span></span> | [<span data-ttu-id="b030d-123">Аддлине</span><span class="sxs-lookup"><span data-stu-id="b030d-123">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="b030d-124">SVG</span><span class="sxs-lookup"><span data-stu-id="b030d-124">SVG</span></span> | [<span data-ttu-id="b030d-125">Аддсвг</span><span class="sxs-lookup"><span data-stu-id="b030d-125">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="b030d-126">Текстовое поле</span><span class="sxs-lookup"><span data-stu-id="b030d-126">Text Box</span></span> | [<span data-ttu-id="b030d-127">Аддтекстбокс</span><span class="sxs-lookup"><span data-stu-id="b030d-127">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="b030d-128">ГеоМетрические фигуры</span><span class="sxs-lookup"><span data-stu-id="b030d-128">Geometric shapes</span></span>

<span data-ttu-id="b030d-129">Создается геометрическая фигура `ShapeCollection.addGeometricShape`.</span><span class="sxs-lookup"><span data-stu-id="b030d-129">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="b030d-130">Этот метод принимает в качестве аргумента перечисление [жеометрикшапетипе](/javascript/api/excel/excel.geometricshapetype) .</span><span class="sxs-lookup"><span data-stu-id="b030d-130">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="b030d-131">В приведенном ниже примере кода создается прямоугольник 150x150-Pixel с именем **"Square"** , который располагается на 100 пикселей от верхней и левой сторон листа.</span><span class="sxs-lookup"><span data-stu-id="b030d-131">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="images"></a><span data-ttu-id="b030d-132">Изображения</span><span class="sxs-lookup"><span data-stu-id="b030d-132">Images</span></span>

<span data-ttu-id="b030d-133">Изображения в формате JPEG, PNG и SVG можно вставить на лист в виде фигур.</span><span class="sxs-lookup"><span data-stu-id="b030d-133">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="b030d-134">`ShapeCollection.addImage` Метод принимает в качестве аргумента строку в кодировке Base64.</span><span class="sxs-lookup"><span data-stu-id="b030d-134">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="b030d-135">Это либо изображение JPEG, либо изображение в формате PNG в виде строки.</span><span class="sxs-lookup"><span data-stu-id="b030d-135">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="b030d-136">`ShapeCollection.addSvg`также принимает в качестве аргумента строку, хотя этот аргумент представляет собой XML, определяющий рисунок.</span><span class="sxs-lookup"><span data-stu-id="b030d-136">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="b030d-137">В следующем примере кода показан файл изображения, загружаемый с помощью [FileReader браузером](https://developer.mozilla.org/docs/Web/API/FileReader) в виде строки.</span><span class="sxs-lookup"><span data-stu-id="b030d-137">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="b030d-138">В строке есть метаданные "base64", которые были удалены перед созданием фигуры.</span><span class="sxs-lookup"><span data-stu-id="b030d-138">The string has the metadata "base64," removed before the shape is created.</span></span>

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = event.target.result.indexOf("base64,");
        var myBase64 = event.target.result.substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a><span data-ttu-id="b030d-139">Lines</span><span class="sxs-lookup"><span data-stu-id="b030d-139">Lines</span></span>

<span data-ttu-id="b030d-140">Создается строка `ShapeCollection.addLine`.</span><span class="sxs-lookup"><span data-stu-id="b030d-140">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="b030d-141">Этот метод требует левое и верхнее поле начальной и конечной точек линии.</span><span class="sxs-lookup"><span data-stu-id="b030d-141">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="b030d-142">Также используется перечисление [коннектортипе](/javascript/api/excel/excel.connectortype) , чтобы указать, как линия контортс между конечными точками.</span><span class="sxs-lookup"><span data-stu-id="b030d-142">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="b030d-143">В примере кода ниже показано, как создать прямую линию на листе.</span><span class="sxs-lookup"><span data-stu-id="b030d-143">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="b030d-144">Линии могут быть связаны с другими объектами Shape.</span><span class="sxs-lookup"><span data-stu-id="b030d-144">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="b030d-145">Методы `connectBeginShape` и `connectEndShape` присоединяют начало и конец линии к фигурам в заданных точках подключения.</span><span class="sxs-lookup"><span data-stu-id="b030d-145">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="b030d-146">Расположение этих точек зависит от формы, но `Shape.connectionSiteCount` можно использовать для того, чтобы надстройка не подключались к точке, которая находится вне границ.</span><span class="sxs-lookup"><span data-stu-id="b030d-146">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="b030d-147">Линия отключается от всех присоединенных фигур с `disconnectBeginShape` помощью `disconnectEndShape` методов и.</span><span class="sxs-lookup"><span data-stu-id="b030d-147">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="b030d-148">В примере кода ниже показано, как подключить строку **"Милине"** к двум фигурам с именами **"лефтшапе"** и **"ригхтшапе"**.</span><span class="sxs-lookup"><span data-stu-id="b030d-148">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-and-resize-shapes"></a><span data-ttu-id="b030d-149">Перемещение и изменение размеров фигур</span><span class="sxs-lookup"><span data-stu-id="b030d-149">Move and resize shapes</span></span>

<span data-ttu-id="b030d-150">Фигуры располагаются поверх листа.</span><span class="sxs-lookup"><span data-stu-id="b030d-150">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="b030d-151">Их размещение определяется свойством `left` и. `top`</span><span class="sxs-lookup"><span data-stu-id="b030d-151">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="b030d-152">Они действуют как поля на соответствующих краях листа, где [0, 0] — верхний левый угол.</span><span class="sxs-lookup"><span data-stu-id="b030d-152">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="b030d-153">Они могут быть установлены напрямую или скорректированы из текущей позиции с помощью методов `incrementLeft` и `incrementTop` .</span><span class="sxs-lookup"><span data-stu-id="b030d-153">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="b030d-154">Размер фигуры, повернутой из положения по умолчанию, также устанавливается таким образом, при этом `rotation` свойство является абсолютной суммой, а `incrementRotation` метод настраивает существующий поворот.</span><span class="sxs-lookup"><span data-stu-id="b030d-154">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="b030d-155">Глубина фигуры относительно других фигур определяется `zorderPosition` свойством.</span><span class="sxs-lookup"><span data-stu-id="b030d-155">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="b030d-156">Это задается с `setZOrder` помощью метода, который принимает [шапезордер](/javascript/api/excel/excel.shapezorder).</span><span class="sxs-lookup"><span data-stu-id="b030d-156">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="b030d-157">`setZOrder`изменяет порядок текущей фигуры относительно других фигур.</span><span class="sxs-lookup"><span data-stu-id="b030d-157">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="b030d-158">У вашей надстройки есть пара параметров для изменения высоты и ширины фигур.</span><span class="sxs-lookup"><span data-stu-id="b030d-158">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="b030d-159">Задание свойства `height` или `width` изменяет указанное измерение без изменения другого измерения.</span><span class="sxs-lookup"><span data-stu-id="b030d-159">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="b030d-160">`scaleHeight` И `scaleWidth` измените соответствующие размеры фигуры относительно текущего или исходного размера (на основе значения предоставленного [шапескалетипе](/javascript/api/excel/excel.shapescaletype)).</span><span class="sxs-lookup"><span data-stu-id="b030d-160">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="b030d-161">Необязательный параметр [шапескалефром](/javascript/api/excel/excel.shapescalefrom) указывает, в каком месте формы масштабируется (верхний левый угол, средний или нижний правый угол).</span><span class="sxs-lookup"><span data-stu-id="b030d-161">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="b030d-162">Если `lockAspectRatio` свойство имеет **значение true**, метод Scale сохраняет текущие пропорции фигуры и настраивает другое измерение.</span><span class="sxs-lookup"><span data-stu-id="b030d-162">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="b030d-163">Прямые изменения свойств `height` и `width` свойств влияют только на это свойство независимо от значения `lockAspectRatio` свойства.</span><span class="sxs-lookup"><span data-stu-id="b030d-163">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="b030d-164">В приведенном ниже примере кода показана фигура, которая масштабируется в 1,25 раз после исходного размера и повернутой 30 градусов.</span><span class="sxs-lookup"><span data-stu-id="b030d-164">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="text-in-shapes"></a><span data-ttu-id="b030d-165">Текст в фигурах</span><span class="sxs-lookup"><span data-stu-id="b030d-165">Text in shapes</span></span>

<span data-ttu-id="b030d-166">ГеоМетрические фигуры могут содержать текст.</span><span class="sxs-lookup"><span data-stu-id="b030d-166">Geometric Shapes can contain text.</span></span> <span data-ttu-id="b030d-167">Фигуры имеют `textFrame` свойство типа [TextFrame](/javascript/api/excel/excel.textframe).</span><span class="sxs-lookup"><span data-stu-id="b030d-167">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="b030d-168">`TextFrame` Объект управляет параметрами отображения текста (например, поля и переполнение текста).</span><span class="sxs-lookup"><span data-stu-id="b030d-168">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="b030d-169">`TextFrame.textRange`— Это объект [TextRange](/javascript/api/excel/excel.textrange) с текстовым контентом и параметрами шрифтов.</span><span class="sxs-lookup"><span data-stu-id="b030d-169">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="b030d-170">В примере кода ниже показано, как создать геометрическую фигуру с именем "Wave" и текстом "текст фигуры".</span><span class="sxs-lookup"><span data-stu-id="b030d-170">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="b030d-171">Он также настраивает цвета фигуры и текста, а также задает горизонтальное выравнивание текста по центру.</span><span class="sxs-lookup"><span data-stu-id="b030d-171">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;
    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");
    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="b030d-172">`addTextBox` Метод `ShapeCollection` , создающий `GeometricShape` тип `Rectangle` с белым фоном и черным текстом.</span><span class="sxs-lookup"><span data-stu-id="b030d-172">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="b030d-173">Это то же самое, что и кнопка Excel в **текстовом поле** на вкладке **Вставка** . `addTextBox` принимает строковый аргумент для задания текста объекта. `TextRange`</span><span class="sxs-lookup"><span data-stu-id="b030d-173">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="b030d-174">В приведенном ниже примере кода показано, как создать текстовое поле с текстом "Hello!".</span><span class="sxs-lookup"><span data-stu-id="b030d-174">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="shape-groups"></a><span data-ttu-id="b030d-175">Группы фигур</span><span class="sxs-lookup"><span data-stu-id="b030d-175">Shape groups</span></span>

<span data-ttu-id="b030d-176">Фигуры можно объединять в группы.</span><span class="sxs-lookup"><span data-stu-id="b030d-176">Shapes can be grouped together.</span></span> <span data-ttu-id="b030d-177">Это позволяет пользователю обрабатывать их как единый объект для позиционирования, изменения размеров и других связанных задач.</span><span class="sxs-lookup"><span data-stu-id="b030d-177">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="b030d-178">[Шапеграуп](/javascript/api/excel/excel.shapegroup) — это тип `Shape`, поэтому надстройка рассматривает группу как отдельную фигуру.</span><span class="sxs-lookup"><span data-stu-id="b030d-178">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="b030d-179">В приведенном ниже примере кода показаны три сгруппированные фигуры.</span><span class="sxs-lookup"><span data-stu-id="b030d-179">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="b030d-180">В следующем примере кода показано, что Группа фигур перемещается вправо на 50 пикселей.</span><span class="sxs-lookup"><span data-stu-id="b030d-180">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var square = shapes.getItem("Square");
    var pentagon = shapes.getItem("Pentagon");
    var octagon = shapes.getItem("Octagon");

    var shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    return context.sync();
}).catch(errorHandlerFunction);

// This sample moves the previously created shape group to the right by 50 pixels.
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shapeGroup = sheet.shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> <span data-ttu-id="b030d-181">Ссылка на отдельные фигуры в группе осуществляется с помощью `ShapeGroup.shapes` свойства, которое относится к типу [граупшапеколлектион](/javascript/api/excel/excel.GroupShapeCollection).</span><span class="sxs-lookup"><span data-stu-id="b030d-181">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="b030d-182">Они больше не доступны через коллекцию фигур листа после группировки.</span><span class="sxs-lookup"><span data-stu-id="b030d-182">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="b030d-183">Например, если лист содержит три фигуры и все они сгруппированы, `shapes.getCount` метод листа возвращает число в 1.</span><span class="sxs-lookup"><span data-stu-id="b030d-183">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="b030d-184">Экспорт фигур в виде изображений</span><span class="sxs-lookup"><span data-stu-id="b030d-184">Export shapes as images</span></span>

<span data-ttu-id="b030d-185">Любой `Shape` объект может быть преобразован в изображение.</span><span class="sxs-lookup"><span data-stu-id="b030d-185">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="b030d-186">[Shape. жетасимаже](/javascript/api/excel/excel.shape#getasimage-format-) возвращает строку в кодировке Base64.</span><span class="sxs-lookup"><span data-stu-id="b030d-186">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="b030d-187">Формат изображения задается в качестве передаваемого перечисление `getAsImage` [PictureFormat](/javascript/api/excel/excel.pictureformat) .</span><span class="sxs-lookup"><span data-stu-id="b030d-187">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

```js
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var shape = sheet.shapes.getItem("Image");
    var stringResult = shape.getAsImage(Excel.PictureFormat.png);

    return context.sync().then(function () {
        console.log(stringResult.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

## <a name="delete-shapes"></a><span data-ttu-id="b030d-188">Удаление фигур</span><span class="sxs-lookup"><span data-stu-id="b030d-188">Delete shapes</span></span>

<span data-ttu-id="b030d-189">Фигуры удаляются из листа с помощью `Shape` `delete` метода объекта.</span><span class="sxs-lookup"><span data-stu-id="b030d-189">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="b030d-190">Другие метаданные не требуются.</span><span class="sxs-lookup"><span data-stu-id="b030d-190">No other metadata is needed.</span></span>

<span data-ttu-id="b030d-191">В примере кода ниже показано, как удалить все фигуры из **миворкшит**.</span><span class="sxs-lookup"><span data-stu-id="b030d-191">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

```js
// This deletes all the shapes from "MyWorksheet".
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("MyWorksheet");
    var shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
            shape.delete()
        });
        return context.sync();
    }).catch(errorHandlerFunction);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="b030d-192">См. также</span><span class="sxs-lookup"><span data-stu-id="b030d-192">See also</span></span>

- [<span data-ttu-id="b030d-193">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b030d-193">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="b030d-194">Работа с диаграммами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="b030d-194">Work with Charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
