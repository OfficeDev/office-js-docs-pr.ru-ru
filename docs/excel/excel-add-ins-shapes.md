---
title: Работать с фигурами с помощью API JavaScript для Excel
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: e4d01c387fff01d68cb26369240a1e06e723a54c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448274"
---
# <a name="work-with-shapes-using-the-excel-javascript-api-preview"></a>Работать с фигурами с помощью API JavaScript для Excel (Предварительная версия)

> [!NOTE]
> API, обсуждаемые в этой статье, в настоящее время доступны только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Excel определяет фигуры как объекты, расположенные в графическом слое Excel. Это означает, что все за прев ячейке ячейка является фигурой. В этой статье описывается, как использовать геометрические фигуры, линии и изображения в сочетании с API [Shape]/жаваскрипт/АПИ/ексцел/ексцел.Шапе) и [ShapeCollection](/javascript/api/excel/excel.shapecollection) . [Диаграммы](/javascript/api/excel/excel.chart) рассматриваются в собственной статье [работать с диаграммами с помощью API JavaScript для Excel]] (Excel-Add-ins-Charts.md)).

## <a name="create-shapes"></a>Создание фигур

Фигуры создаются и хранятся в коллекции фигур листа (`Worksheet.shapes`). `ShapeCollection`в этой `.add*` цели есть несколько способов. Все фигуры имеют имена и идентификаторы, созданные для них при добавлении в коллекцию. Это свойства `name` и `id` , соответственно, свойства. `name`может быть задано надстройкой для упрощения поиска с помощью `ShapeCollection.getItem(name)` метода.

С помощью связанного метода добавляются следующие типы фигур:

| Shape | Добавление метода | Подпись |
|-------|------------|-----------|
| ГеоМетрическая фигура | [Адджеометрикшапе](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Изображение (JPEG или PNG) | [Аддимаже](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [Аддлине](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [Аддсвг](/javascript/api/excel/excel.shapecollection#addsvg-xml-) | `addSvg(xml: string): Excel.Shape` |
| Текстовое поле | [Аддтекстбокс](/javascript/api/excel/excel.shapecollection#addtextbox-text-) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>ГеоМетрические фигуры

Создается геометрическая фигура `ShapeCollection.addGeometricShape`. Этот метод принимает в качестве аргумента перечисление [жеометрикшапетипе](/javascript/api/excel/excel.geometricshapetype) .

В приведенном ниже примере кода создается прямоугольник 150x150-Pixel с именем **"Square"** , который располагается на 100 пикселей от верхней и левой сторон листа.

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

### <a name="images"></a>Изображения

Изображения в формате JPEG, PNG и SVG можно вставить на лист в виде фигур. `ShapeCollection.addImage` Метод принимает в качестве аргумента строку в кодировке Base64. Это либо изображение JPEG, либо изображение в формате PNG в виде строки. `ShapeCollection.addSvg`также принимает в качестве аргумента строку, хотя этот аргумент представляет собой XML, определяющий рисунок.

В следующем примере кода показан файл изображения, загружаемый с помощью [FileReader браузером](https://developer.mozilla.org/docs/Web/API/FileReader) в виде строки. В строке есть метаданные "base64", которые были удалены перед созданием фигуры.

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

### <a name="lines"></a>Lines

Создается строка `ShapeCollection.addLine`. Этот метод требует левое и верхнее поле начальной и конечной точек линии. Также используется перечисление [коннектортипе](/javascript/api/excel/excel.connectortype) , чтобы указать, как линия контортс между конечными точками. В примере кода ниже показано, как создать прямую линию на листе.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet
Excel.run(function (context) {
    var shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    var line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    return context.sync();
}).catch(errorHandlerFunction);
```

Линии могут быть связаны с другими объектами Shape. Методы `connectBeginShape` и `connectEndShape` присоединяют начало и конец линии к фигурам в заданных точках подключения. Расположение этих точек зависит от формы, но `Shape.connectionSiteCount` можно использовать для того, чтобы надстройка не подключались к точке, которая находится вне границ. Линия отключается от всех присоединенных фигур с `disconnectBeginShape` помощью `disconnectEndShape` методов и.

В примере кода ниже показано, как подключить строку **"Милине"** к двум фигурам с именами **"лефтшапе"** и **"ригхтшапе"**.

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

## <a name="move-and-resize-shapes"></a>Перемещение и изменение размеров фигур

Фигуры располагаются поверх листа. Их размещение определяется свойством `left` и. `top` Они действуют как поля на соответствующих краях листа, где [0, 0] — верхний левый угол. Они могут быть установлены напрямую или скорректированы из текущей позиции с помощью методов `incrementLeft` и `incrementTop` . Размер фигуры, повернутой из положения по умолчанию, также устанавливается таким образом, при этом `rotation` свойство является абсолютной суммой, а `incrementRotation` метод настраивает существующий поворот.

Глубина фигуры относительно других фигур определяется `zorderPosition` свойством. Это задается с `setZOrder` помощью метода, который принимает [шапезордер](/javascript/api/excel/excel.shapezorder). `setZOrder`изменяет порядок текущей фигуры относительно других фигур.

У вашей надстройки есть пара параметров для изменения высоты и ширины фигур. Задание свойства `height` или `width` изменяет указанное измерение без изменения другого измерения. `scaleHeight` И `scaleWidth` измените соответствующие размеры фигуры относительно текущего или исходного размера (на основе значения предоставленного [шапескалетипе](/javascript/api/excel/excel.shapescaletype)). Необязательный параметр [шапескалефром](/javascript/api/excel/excel.shapescalefrom) указывает, в каком месте формы масштабируется (верхний левый угол, средний или нижний правый угол). Если `lockAspectRatio` свойство имеет **значение true**, метод Scale сохраняет текущие пропорции фигуры и настраивает другое измерение.

> [!NOTE]
> Прямые изменения свойств `height` и `width` свойств влияют только на это свойство независимо от значения `lockAspectRatio` свойства.

В приведенном ниже примере кода показана фигура, которая масштабируется в 1,25 раз после исходного размера и повернутой 30 градусов.

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

## <a name="text-in-shapes"></a>Текст в фигурах

ГеоМетрические фигуры могут содержать текст. Фигуры имеют `textFrame` свойство типа [TextFrame](/javascript/api/excel/excel.textframe). `TextFrame` Объект управляет параметрами отображения текста (например, поля и переполнение текста). `TextFrame.textRange`— Это объект [TextRange](/javascript/api/excel/excel.textrange) с текстовым контентом и параметрами шрифтов.

В примере кода ниже показано, как создать геометрическую фигуру с именем "Wave" и текстом "текст фигуры". Он также настраивает цвета фигуры и текста, а также задает горизонтальное выравнивание текста по центру.

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

`addTextBox` Метод `ShapeCollection` , создающий `GeometricShape` тип `Rectangle` с белым фоном и черным текстом. Это то же самое, что и кнопка Excel в **текстовом поле** на вкладке **Вставка** . `addTextBox` принимает строковый аргумент для задания текста объекта. `TextRange`

В приведенном ниже примере кода показано, как создать текстовое поле с текстом "Hello!".

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

## <a name="shape-groups"></a>Группы фигур

Фигуры можно объединять в группы. Это позволяет пользователю обрабатывать их как единый объект для позиционирования, изменения размеров и других связанных задач. [Шапеграуп](/javascript/api/excel/excel.shapegroup) — это тип `Shape`, поэтому надстройка рассматривает группу как отдельную фигуру.

В приведенном ниже примере кода показаны три сгруппированные фигуры. В следующем примере кода показано, что Группа фигур перемещается вправо на 50 пикселей.

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
> Ссылка на отдельные фигуры в группе осуществляется с помощью `ShapeGroup.shapes` свойства, которое относится к типу [граупшапеколлектион](/javascript/api/excel/excel.GroupShapeCollection). Они больше не доступны через коллекцию фигур листа после группировки. Например, если лист содержит три фигуры и все они сгруппированы, `shapes.getCount` метод листа возвращает число в 1.

## <a name="export-shapes-as-images"></a>Экспорт фигур в виде изображений

Любой `Shape` объект может быть преобразован в изображение. [Shape. жетасимаже](/javascript/api/excel/excel.shape#getasimage-format-) возвращает строку в кодировке Base64. Формат изображения задается в качестве передаваемого перечисление `getAsImage` [PictureFormat](/javascript/api/excel/excel.pictureformat) .

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

## <a name="delete-shapes"></a>Удаление фигур

Фигуры удаляются из листа с помощью `Shape` `delete` метода объекта. Другие метаданные не требуются.

В примере кода ниже показано, как удалить все фигуры из **миворкшит**.

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

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Работа с диаграммами с использованием API JavaScript для Excel](excel-add-ins-charts.md)
