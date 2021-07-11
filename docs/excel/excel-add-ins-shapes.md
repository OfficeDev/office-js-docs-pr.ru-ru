---
title: Работа с фигурами с Excel API JavaScript
description: Узнайте, Excel определяет фигуры как любой объект, который находится на уровне рисования Excel.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: eeb6a1f76c839e4b550662b28b717bfd1bcca4e8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349450"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a><span data-ttu-id="8daa8-103">Работа с фигурами с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="8daa8-103">Work with shapes using the Excel JavaScript API</span></span>

<span data-ttu-id="8daa8-104">Excel определяет фигуры как любой объект, который находится на уровне рисования Excel.</span><span class="sxs-lookup"><span data-stu-id="8daa8-104">Excel defines shapes as any object that sits on the drawing layer of Excel.</span></span> <span data-ttu-id="8daa8-105">Это означает, что все, что находится за пределами ячейки, — это фигура.</span><span class="sxs-lookup"><span data-stu-id="8daa8-105">That means anything outside of a cell is a shape.</span></span> <span data-ttu-id="8daa8-106">В этой статье описывается использование геометрических фигур, линий и изображений в сочетании с API [Shape](/javascript/api/excel/excel.shape) и [ShapeCollection.](/javascript/api/excel/excel.shapecollection)</span><span class="sxs-lookup"><span data-stu-id="8daa8-106">This article describes how to use geometric shapes, lines, and images in conjunction with the [Shape](/javascript/api/excel/excel.shape) and [ShapeCollection](/javascript/api/excel/excel.shapecollection) APIs.</span></span> <span data-ttu-id="8daa8-107">[Диаграммы](/javascript/api/excel/excel.chart) охватываются в своей статье , Работа с диаграммами с помощью [Excel API JavaScript](excel-add-ins-charts.md).</span><span class="sxs-lookup"><span data-stu-id="8daa8-107">[Charts](/javascript/api/excel/excel.chart) are covered in their own article, [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

<span data-ttu-id="8daa8-108">На следующем изображении показаны фигуры, которые образуют термометр.</span><span class="sxs-lookup"><span data-stu-id="8daa8-108">The following image shows shapes which form a thermometer.</span></span>
<span data-ttu-id="8daa8-109">![Изображение термометра, выполненного в виде Excel формы.](../images/excel-shapes.png)</span><span class="sxs-lookup"><span data-stu-id="8daa8-109">![Image of a thermometer made as an Excel shape.](../images/excel-shapes.png)</span></span>

## <a name="create-shapes"></a><span data-ttu-id="8daa8-110">Создание фигур</span><span class="sxs-lookup"><span data-stu-id="8daa8-110">Create shapes</span></span>

<span data-ttu-id="8daa8-111">Формы создаются с помощью и хранятся в коллекции фигуры таблицы ( `Worksheet.shapes` ).</span><span class="sxs-lookup"><span data-stu-id="8daa8-111">Shapes are created through and stored in a worksheet's shape collection (`Worksheet.shapes`).</span></span> <span data-ttu-id="8daa8-112">`ShapeCollection` имеет несколько `.add*` методов для этой цели.</span><span class="sxs-lookup"><span data-stu-id="8daa8-112">`ShapeCollection` has several `.add*` methods for this purpose.</span></span> <span data-ttu-id="8daa8-113">Все фигуры имеют имена и ИД, созданные для них при добавлении в коллекцию.</span><span class="sxs-lookup"><span data-stu-id="8daa8-113">All shapes have names and IDs generated for them when they are added to the collection.</span></span> <span data-ttu-id="8daa8-114">Это и `name` `id` свойства, соответственно.</span><span class="sxs-lookup"><span data-stu-id="8daa8-114">These are the `name` and `id` properties, respectively.</span></span> <span data-ttu-id="8daa8-115">`name` может быть установлено вашей надстройки для легкого получения с помощью `ShapeCollection.getItem(name)` метода.</span><span class="sxs-lookup"><span data-stu-id="8daa8-115">`name` can be set by your add-in for easy retrieval with the `ShapeCollection.getItem(name)` method.</span></span>

<span data-ttu-id="8daa8-116">Следующие типы фигур добавляются с помощью связанного метода.</span><span class="sxs-lookup"><span data-stu-id="8daa8-116">The following types of shapes are added using the associated method.</span></span>

| <span data-ttu-id="8daa8-117">Shape</span><span class="sxs-lookup"><span data-stu-id="8daa8-117">Shape</span></span> | <span data-ttu-id="8daa8-118">Добавление метода</span><span class="sxs-lookup"><span data-stu-id="8daa8-118">Add Method</span></span> | <span data-ttu-id="8daa8-119">Подпись</span><span class="sxs-lookup"><span data-stu-id="8daa8-119">Signature</span></span> |
|-------|------------|-----------|
| <span data-ttu-id="8daa8-120">Геометрическая фигура</span><span class="sxs-lookup"><span data-stu-id="8daa8-120">Geometric Shape</span></span> | [<span data-ttu-id="8daa8-121">addGeometricShape</span><span class="sxs-lookup"><span data-stu-id="8daa8-121">addGeometricShape</span></span>](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| <span data-ttu-id="8daa8-122">Изображение (JPEG или PNG)</span><span class="sxs-lookup"><span data-stu-id="8daa8-122">Image (either JPEG or PNG)</span></span> | [<span data-ttu-id="8daa8-123">addImage</span><span class="sxs-lookup"><span data-stu-id="8daa8-123">addImage</span></span>](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| <span data-ttu-id="8daa8-124">Line</span><span class="sxs-lookup"><span data-stu-id="8daa8-124">Line</span></span> | [<span data-ttu-id="8daa8-125">addLine</span><span class="sxs-lookup"><span data-stu-id="8daa8-125">addLine</span></span>](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| <span data-ttu-id="8daa8-126">SVG</span><span class="sxs-lookup"><span data-stu-id="8daa8-126">SVG</span></span> | [<span data-ttu-id="8daa8-127">addSvg</span><span class="sxs-lookup"><span data-stu-id="8daa8-127">addSvg</span></span>](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| <span data-ttu-id="8daa8-128">Текстовое поле</span><span class="sxs-lookup"><span data-stu-id="8daa8-128">Text Box</span></span> | [<span data-ttu-id="8daa8-129">addTextBox</span><span class="sxs-lookup"><span data-stu-id="8daa8-129">addTextBox</span></span>](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a><span data-ttu-id="8daa8-130">Геометрические фигуры</span><span class="sxs-lookup"><span data-stu-id="8daa8-130">Geometric shapes</span></span>

<span data-ttu-id="8daa8-131">Геометрическая фигура создается с `ShapeCollection.addGeometricShape` помощью .</span><span class="sxs-lookup"><span data-stu-id="8daa8-131">A geometric shape is created with `ShapeCollection.addGeometricShape`.</span></span> <span data-ttu-id="8daa8-132">Этот метод принимает в качестве аргумента enum [GeometricShapeType.](/javascript/api/excel/excel.geometricshapetype)</span><span class="sxs-lookup"><span data-stu-id="8daa8-132">That method takes a [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) enum as an argument.</span></span>

<span data-ttu-id="8daa8-133">В следующем примере кода создается прямоугольник размером 150x150 пикселей с именем **"Square",** который находится на 100 пикселей с верхней и левой сторон таблицы.</span><span class="sxs-lookup"><span data-stu-id="8daa8-133">The following code sample creates a 150x150-pixel rectangle named **"Square"** that is positioned 100 pixels from the top and left sides of the worksheet.</span></span>

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

### <a name="images"></a><span data-ttu-id="8daa8-134">изображения;</span><span class="sxs-lookup"><span data-stu-id="8daa8-134">Images</span></span>

<span data-ttu-id="8daa8-135">Изображения JPEG, PNG и SVG можно вставить в таблицу в форме фигур.</span><span class="sxs-lookup"><span data-stu-id="8daa8-135">JPEG, PNG, and SVG images can be inserted into a worksheet as shapes.</span></span> <span data-ttu-id="8daa8-136">В качестве аргумента метод принимает строку с кодом `ShapeCollection.addImage` base64.</span><span class="sxs-lookup"><span data-stu-id="8daa8-136">The `ShapeCollection.addImage` method takes a base64-encoded string as an argument.</span></span> <span data-ttu-id="8daa8-137">Это либо образ JPEG или PNG в строковом виде.</span><span class="sxs-lookup"><span data-stu-id="8daa8-137">This is either a JPEG or PNG image in string form.</span></span> <span data-ttu-id="8daa8-138">`ShapeCollection.addSvg` также принимает строку, хотя этот аргумент XML, который определяет графику.</span><span class="sxs-lookup"><span data-stu-id="8daa8-138">`ShapeCollection.addSvg` also takes in a string, though this argument is XML that defines the graphic.</span></span>

<span data-ttu-id="8daa8-139">В следующем примере кода показан файл изображения, загружаемый [файлом FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) в качестве строки.</span><span class="sxs-lookup"><span data-stu-id="8daa8-139">The following code sample shows an image file being loaded by a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) as a string.</span></span> <span data-ttu-id="8daa8-140">Строка имеет метаданные "base64", удалены до создания фигуры.</span><span class="sxs-lookup"><span data-stu-id="8daa8-140">The string has the metadata "base64," removed before the shape is created.</span></span>

```js
// This sample creates an image as a Shape object in the worksheet.
var myFile = document.getElementById("selectedFile");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        var startIndex = reader.result.toString().indexOf("base64,");
        var myBase64 = reader.result.toString().substr(startIndex + 7);
        var sheet = context.workbook.worksheets.getItem("MyWorksheet");
        var image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a><span data-ttu-id="8daa8-141">Lines</span><span class="sxs-lookup"><span data-stu-id="8daa8-141">Lines</span></span>

<span data-ttu-id="8daa8-142">Строка создается с `ShapeCollection.addLine` помощью .</span><span class="sxs-lookup"><span data-stu-id="8daa8-142">A line is created with `ShapeCollection.addLine`.</span></span> <span data-ttu-id="8daa8-143">Для этого метода необходимы левые и верхние поля точки начала и конца строки.</span><span class="sxs-lookup"><span data-stu-id="8daa8-143">That method needs the left and top margins of the line's start and end points.</span></span> <span data-ttu-id="8daa8-144">Кроме того, для указания того, как соединителю строки между конечными точками, необходимо также вводить в себя enum [ConnectorType.](/javascript/api/excel/excel.connectortype)</span><span class="sxs-lookup"><span data-stu-id="8daa8-144">It also takes a [ConnectorType](/javascript/api/excel/excel.connectortype) enum to specify how the line contorts between endpoints.</span></span> <span data-ttu-id="8daa8-145">В следующем примере кода создается прямая линия на таблице.</span><span class="sxs-lookup"><span data-stu-id="8daa8-145">The following code sample creates a straight line on the worksheet.</span></span>

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="8daa8-146">Строки можно подключать к другим объектам Shape.</span><span class="sxs-lookup"><span data-stu-id="8daa8-146">Lines can be connected to other Shape objects.</span></span> <span data-ttu-id="8daa8-147">Методы и начало и окончание строки прикрепляются к `connectBeginShape` `connectEndShape` фигурам в указанных точках подключения.</span><span class="sxs-lookup"><span data-stu-id="8daa8-147">The `connectBeginShape` and `connectEndShape` methods attach the beginning and ending of a line to shapes at the specified connection points.</span></span> <span data-ttu-id="8daa8-148">Расположение этих точек зависит от формы, но его можно использовать для обеспечения того, чтобы надстройка не подключалась к точке, не связанной `Shape.connectionSiteCount` с этим.</span><span class="sxs-lookup"><span data-stu-id="8daa8-148">The locations of these points vary by shape, but the `Shape.connectionSiteCount` can be used to ensure your add-in does not connect to a point that's out-of-bounds.</span></span> <span data-ttu-id="8daa8-149">Строка отключена от любых присоединенных фигур с помощью `disconnectBeginShape` этих и `disconnectEndShape` методов.</span><span class="sxs-lookup"><span data-stu-id="8daa8-149">A line is disconnected from any attached shapes using the `disconnectBeginShape` and `disconnectEndShape` methods.</span></span>

<span data-ttu-id="8daa8-150">В следующем примере кода **строка "MyLine"** соединяется с двумя фигурами с именем **"LeftShape"** **и "RightShape".**</span><span class="sxs-lookup"><span data-stu-id="8daa8-150">The following code sample connects the **"MyLine"** line to two shapes named **"LeftShape"** and **"RightShape"**.</span></span>

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

## <a name="move-and-resize-shapes"></a><span data-ttu-id="8daa8-151">Перемещение и размер фигур</span><span class="sxs-lookup"><span data-stu-id="8daa8-151">Move and resize shapes</span></span>

<span data-ttu-id="8daa8-152">Фигуры сидят на вершине таблицы.</span><span class="sxs-lookup"><span data-stu-id="8daa8-152">Shapes sit on top of the worksheet.</span></span> <span data-ttu-id="8daa8-153">Их размещение определяется свойством `left` `top` и свойством.</span><span class="sxs-lookup"><span data-stu-id="8daa8-153">Their placement is defined by the `left` and `top` property.</span></span> <span data-ttu-id="8daa8-154">Они действуют как поля от соответствующих краев таблицы, а [0, 0] — верхний левый угол.</span><span class="sxs-lookup"><span data-stu-id="8daa8-154">These act as margins from worksheet's respective edges, with [0, 0] being the upper-left corner.</span></span> <span data-ttu-id="8daa8-155">Они могут быть установлены непосредственно или скорректированы с текущей позиции с помощью `incrementLeft` методов и `incrementTop` методов.</span><span class="sxs-lookup"><span data-stu-id="8daa8-155">These can either be set directly or adjusted from their current position with the `incrementLeft` and `incrementTop` methods.</span></span> <span data-ttu-id="8daa8-156">Таким образом устанавливается также размер поворота фигуры из положения по умолчанию, при этом свойство является абсолютным количеством и методом, корректющим `rotation` `incrementRotation` существующее вращение.</span><span class="sxs-lookup"><span data-stu-id="8daa8-156">How much a shape is rotated from the default position is also established in this manner, with the `rotation` property being the absolute amount and the `incrementRotation` method adjusting the existing rotation.</span></span>

<span data-ttu-id="8daa8-157">Глубина фигуры по отношению к другим фигурам определяется `zorderPosition` свойством.</span><span class="sxs-lookup"><span data-stu-id="8daa8-157">A shape's depth relative to other shapes is defined by the `zorderPosition` property.</span></span> <span data-ttu-id="8daa8-158">Это устанавливается с помощью `setZOrder` метода, который принимает [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span><span class="sxs-lookup"><span data-stu-id="8daa8-158">This is set using the `setZOrder` method, which takes a [ShapeZOrder](/javascript/api/excel/excel.shapezorder).</span></span> <span data-ttu-id="8daa8-159">`setZOrder` регулирует порядок текущей фигуры по отношению к другим фигурам.</span><span class="sxs-lookup"><span data-stu-id="8daa8-159">`setZOrder` adjusts the ordering of the current shape relative to the other shapes.</span></span>

<span data-ttu-id="8daa8-160">Надстройка имеет несколько вариантов изменения высоты и ширины фигур.</span><span class="sxs-lookup"><span data-stu-id="8daa8-160">Your add-in has a couple options for changing the height and width of shapes.</span></span> <span data-ttu-id="8daa8-161">Параметр или `height` свойство `width` изменяет указанное измерение без изменения другого измерения.</span><span class="sxs-lookup"><span data-stu-id="8daa8-161">Setting either the `height` or `width` property changes the specified dimension without changing the other dimension.</span></span> <span data-ttu-id="8daa8-162">Соответствующие размеры фигуры по отношению к текущему или исходному размеру (в зависимости от значения предоставленного `scaleHeight` `scaleWidth` [ShapeScaleType)](/javascript/api/excel/excel.shapescaletype)и настройки.</span><span class="sxs-lookup"><span data-stu-id="8daa8-162">The `scaleHeight` and `scaleWidth` adjust the shape's respective dimensions relative to either the current or original size (based on the value of the provided [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)).</span></span> <span data-ttu-id="8daa8-163">Необязательный параметр [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) указывает, откуда масштабирует фигуру (верхний левый угол, средний или нижний правый угол).</span><span class="sxs-lookup"><span data-stu-id="8daa8-163">An optional [ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) parameter specifies from where the shape scales (top-left corner, middle, or bottom-right corner).</span></span> <span data-ttu-id="8daa8-164">Если свойство верно, методы масштабирования поддерживают текущее отношение аспектов фигуры, а также корректируют `lockAspectRatio` другое измерение. </span><span class="sxs-lookup"><span data-stu-id="8daa8-164">If the `lockAspectRatio` property is **true**, the scale methods maintain the shape's current aspect ratio by also adjusting the other dimension.</span></span>

> [!NOTE]
> <span data-ttu-id="8daa8-165">Прямые изменения свойств и свойств влияют только на это свойство, независимо от `height` `width` значения `lockAspectRatio` свойства.</span><span class="sxs-lookup"><span data-stu-id="8daa8-165">Direct changes to the `height` and `width` properties only affect that property, regardless of the `lockAspectRatio` property's value.</span></span>

<span data-ttu-id="8daa8-166">В следующем примере кода показана фигура, масштабироваться в 1,25 раза и вращаться на 30 градусов.</span><span class="sxs-lookup"><span data-stu-id="8daa8-166">The following code sample shows a shape being scaled to 1.25 times its original size and rotated 30 degrees.</span></span>

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

## <a name="text-in-shapes"></a><span data-ttu-id="8daa8-167">Текст в фигурах</span><span class="sxs-lookup"><span data-stu-id="8daa8-167">Text in shapes</span></span>

<span data-ttu-id="8daa8-168">Геометрические фигуры могут содержать текст.</span><span class="sxs-lookup"><span data-stu-id="8daa8-168">Geometric Shapes can contain text.</span></span> <span data-ttu-id="8daa8-169">Фигуры имеют `textFrame` свойство типа [TextFrame](/javascript/api/excel/excel.textframe).</span><span class="sxs-lookup"><span data-stu-id="8daa8-169">Shapes have a `textFrame` property of type [TextFrame](/javascript/api/excel/excel.textframe).</span></span> <span data-ttu-id="8daa8-170">Объект `TextFrame` управляет вариантами отображения текста (например, полями и переполнением текста).</span><span class="sxs-lookup"><span data-stu-id="8daa8-170">The `TextFrame` object manages the text display options (such as margins and text overflow).</span></span> <span data-ttu-id="8daa8-171">`TextFrame.textRange` — объект [TextRange](/javascript/api/excel/excel.textrange) с текстовым контентом и настройками шрифтов.</span><span class="sxs-lookup"><span data-stu-id="8daa8-171">`TextFrame.textRange` is a [TextRange](/javascript/api/excel/excel.textrange) object with the text content and font settings.</span></span>

<span data-ttu-id="8daa8-172">В следующем примере кода создается геометрическая фигура с именем "Wave" с текстом "Shape Text".</span><span class="sxs-lookup"><span data-stu-id="8daa8-172">The following code sample creates a geometric shape named "Wave" with the text "Shape Text".</span></span> <span data-ttu-id="8daa8-173">Он также регулирует форму и текстовые цвета, а также задает горизонтальное выравнивание текста в центре.</span><span class="sxs-lookup"><span data-stu-id="8daa8-173">It also adjusts the shape and text colors, as well as sets the text's horizontal alignment to the center.</span></span>

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

<span data-ttu-id="8daa8-174">Метод создает тип с белым фоном `addTextBox` `ShapeCollection` и черным `GeometricShape` `Rectangle` текстом.</span><span class="sxs-lookup"><span data-stu-id="8daa8-174">The `addTextBox` method of `ShapeCollection` creates a `GeometricShape` of type `Rectangle` with a white background and black text.</span></span> <span data-ttu-id="8daa8-175">Это то же самое, что и Excel **на** вкладке  Вставка. `addTextBox` принимает аргумент строки, чтобы установить текст `TextRange` .</span><span class="sxs-lookup"><span data-stu-id="8daa8-175">This is the same as what is created by Excel's **Text Box** button on the **Insert** tab. `addTextBox` takes a string argument to set the text of the `TextRange`.</span></span>

<span data-ttu-id="8daa8-176">В следующем примере кода показано создание текстового окна с текстом "Hello!".</span><span class="sxs-lookup"><span data-stu-id="8daa8-176">The following code sample shows the creation of a text box with the text "Hello!".</span></span>

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

## <a name="shape-groups"></a><span data-ttu-id="8daa8-177">Группы формы</span><span class="sxs-lookup"><span data-stu-id="8daa8-177">Shape groups</span></span>

<span data-ttu-id="8daa8-178">Фигуры можно сгруппить вместе.</span><span class="sxs-lookup"><span data-stu-id="8daa8-178">Shapes can be grouped together.</span></span> <span data-ttu-id="8daa8-179">Это позволяет пользователю рассматривать их как единую сущность для позиционирования, размеров и других связанных задач.</span><span class="sxs-lookup"><span data-stu-id="8daa8-179">This allows a user to treat them as a single entity for positioning, sizing, and other related tasks.</span></span> <span data-ttu-id="8daa8-180">[ShapeGroup](/javascript/api/excel/excel.shapegroup) — это тип, поэтому ваша надстройка рассматривает группу `Shape` как единую фигуру.</span><span class="sxs-lookup"><span data-stu-id="8daa8-180">A [ShapeGroup](/javascript/api/excel/excel.shapegroup) is a type of `Shape`, so your add-in treats the group as a single shape.</span></span>

<span data-ttu-id="8daa8-181">В следующем примере кода показаны три фигуры, сгруппироваться вместе.</span><span class="sxs-lookup"><span data-stu-id="8daa8-181">The following code sample shows three shapes being grouped together.</span></span> <span data-ttu-id="8daa8-182">В последующем примере кода показано, что группа фигур перемещается в нужные 50 пикселей.</span><span class="sxs-lookup"><span data-stu-id="8daa8-182">The subsequent code sample shows that shape group being moved to the right 50 pixels.</span></span>

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
> <span data-ttu-id="8daa8-183">Отдельные фигуры в группе ссылаются через свойство, которое имеет `ShapeGroup.shapes` тип [GroupShapeCollection.](/javascript/api/excel/excel.GroupShapeCollection)</span><span class="sxs-lookup"><span data-stu-id="8daa8-183">Individual shapes within the group are referenced through the `ShapeGroup.shapes` property, which is of type [GroupShapeCollection](/javascript/api/excel/excel.GroupShapeCollection).</span></span> <span data-ttu-id="8daa8-184">Они больше не доступны в коллекции фигуры таблицы после сгруппии.</span><span class="sxs-lookup"><span data-stu-id="8daa8-184">They are no longer accessible through the worksheet's shape collection after being grouped.</span></span> <span data-ttu-id="8daa8-185">Например, если ваш таблица имеет три фигуры и все они сгруппировалися, метод таблицы возвращает количество `shapes.getCount` 1.</span><span class="sxs-lookup"><span data-stu-id="8daa8-185">As an example, if your worksheet had three shapes and they were all grouped together, the worksheet's `shapes.getCount` method would return a count of 1.</span></span>

## <a name="export-shapes-as-images"></a><span data-ttu-id="8daa8-186">Экспорт фигур в качестве изображений</span><span class="sxs-lookup"><span data-stu-id="8daa8-186">Export shapes as images</span></span>

<span data-ttu-id="8daa8-187">Любой `Shape` объект можно преобразовать в изображение.</span><span class="sxs-lookup"><span data-stu-id="8daa8-187">Any `Shape` object can be converted to an image.</span></span> <span data-ttu-id="8daa8-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) возвращает строку base64-encoded.</span><span class="sxs-lookup"><span data-stu-id="8daa8-188">[Shape.getAsImage](/javascript/api/excel/excel.shape#getasimage-format-) returns base64-encoded string.</span></span> <span data-ttu-id="8daa8-189">Формат изображения указывается как переоформаемая в [](/javascript/api/excel/excel.pictureformat) `getAsImage` .</span><span class="sxs-lookup"><span data-stu-id="8daa8-189">The image's format is specified as a [PictureFormat](/javascript/api/excel/excel.pictureformat) enum passed to `getAsImage`.</span></span>

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

## <a name="delete-shapes"></a><span data-ttu-id="8daa8-190">Удаление фигур</span><span class="sxs-lookup"><span data-stu-id="8daa8-190">Delete shapes</span></span>

<span data-ttu-id="8daa8-191">Фигуры удаляются из таблицы методом `Shape` `delete` объекта.</span><span class="sxs-lookup"><span data-stu-id="8daa8-191">Shapes are removed from the worksheet with the `Shape` object's `delete` method.</span></span> <span data-ttu-id="8daa8-192">Другие метаданные не нужны.</span><span class="sxs-lookup"><span data-stu-id="8daa8-192">No other metadata is needed.</span></span>

<span data-ttu-id="8daa8-193">В следующем примере кода удаляются все фигуры **из MyWorksheet.**</span><span class="sxs-lookup"><span data-stu-id="8daa8-193">The following code sample deletes all the shapes from **MyWorksheet**.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="8daa8-194">См. также</span><span class="sxs-lookup"><span data-stu-id="8daa8-194">See also</span></span>

- [<span data-ttu-id="8daa8-195">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="8daa8-195">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="8daa8-196">Работа с диаграммами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="8daa8-196">Work with charts using the Excel JavaScript API</span></span>](excel-add-ins-charts.md)
