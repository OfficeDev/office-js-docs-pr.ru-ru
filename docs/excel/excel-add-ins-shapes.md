---
title: Работа с фигурами с помощью API JavaScript для Excel
description: Узнайте, как Excel определяет фигуры как любой объект, который находится на слое рисования Excel.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 507ae05b570e7eef4f3bf5560ca47c1bfbd40f9f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889599"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Работа с фигурами с помощью API JavaScript для Excel

Excel определяет фигуры как любой объект, который находится на слое документа Excel. Это означает, что все, что находится за пределами ячейки, является фигурой. В этой статье описывается, как использовать геометрические фигуры, линии и изображения в сочетании с [API-интерфейсами Shape и ShapeCollection](/javascript/api/excel/excel.shapecollection).[](/javascript/api/excel/excel.shape) [Диаграммы](/javascript/api/excel/excel.chart) рассматриваются в собственной статье " [Работа с диаграммами с помощью API JavaScript для Excel"](excel-add-ins-charts.md).

На следующем рисунке показаны фигуры, которые образуют термометр.
![Изображение термометра, выполненного в виде фигуры Excel.](../images/excel-shapes.png)

## <a name="create-shapes"></a>Создание фигур

Фигуры создаются и хранятся в коллекции фигур листа ().`Worksheet.shapes` `ShapeCollection` имеет несколько `.add*` методов для этой цели. Все фигуры имеют имена и идентификаторы, созданные для них при добавлении в коллекцию. Это свойства `name` `id` и свойства соответственно. `name` можно задать надстройке для простого извлечения с помощью `ShapeCollection.getItem(name)` метода.

Следующие типы фигур добавляются с помощью связанного метода.

| Shape | Добавление метода | Подпись |
|-------|------------|-----------|
| Геометрическая фигура | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Изображение (JPEG или PNG) | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| Текстовое поле | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>Геометрические фигуры

Геометрическая фигура создается с помощью `ShapeCollection.addGeometricShape`. Этот метод принимает [перечисление GeometrShapeType](/javascript/api/excel/excel.geometricshapetype) в качестве аргумента.

В следующем примере кода создается прямоугольник размером 150 x 150 пикселей с именем **Square** , который располагается в 100 пикселей от верхней и левой сторон листа.

```js
// This sample creates a rectangle positioned 100 pixels from the top and left sides
// of the worksheet and is 150x150 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;

    let rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";

    await context.sync();
});
```

### <a name="images"></a>изображения;

Изображения JPEG, PNG и SVG можно вставлять на лист в виде фигур. Метод `ShapeCollection.addImage` принимает строку в кодировке base64 в качестве аргумента. Это изображение JPEG или PNG в строковом формате. `ShapeCollection.addSvg` также принимает строку, хотя этот аргумент является XML, который определяет рисунок.

В следующем примере кода показан файл изображения, загружаемый [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) в виде строки. Строка содержит метаданные base64, удаленные до создания фигуры.

```js
// This sample creates an image as a Shape object in the worksheet.
let myFile = document.getElementById("selectedFile");
let reader = new FileReader();

reader.onload = (event) => {
    Excel.run(function (context) {
        let startIndex = reader.result.toString().indexOf("base64,");
        let myBase64 = reader.result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getItem("MyWorksheet");
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";
        return context.sync();
    }).catch(errorHandlerFunction);
};

// Read in the image file as a data URL.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="lines"></a>Lines

Строка создается с помощью `ShapeCollection.addLine`. Для этого метода требуются левые и верхние поля начальной и конечной точек строки. Он также принимает [перечисление ConnectorType](/javascript/api/excel/excel.connectortype) , чтобы указать, как линия перемыкается между конечными точками. В следующем примере кода создается прямая линия на листе.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

Линии могут быть соединены с другими объектами Shape. И `connectBeginShape` методы `connectEndShape` присоединяют начало и конец строки к фигурам в указанных точках соединения. Расположения этих точек зависят от фигуры, `Shape.connectionSiteCount` но их можно использовать для того, чтобы надстройка не подключается к точке, которая не является границей. Линия отсоединяется от всех присоединенных фигур с помощью методов `disconnectBeginShape` и методов `disconnectEndShape` .

В следующем примере кода строка **MyLine** соединяется с двумя фигурами с именами **LeftShape** и **RightShape**.

```js
// This sample connects a line between two shapes at connection points '0' and '3'.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.getItem("MyLine").line;
    line.connectBeginShape(shapes.getItem("LeftShape"), 0);
    line.connectEndShape(shapes.getItem("RightShape"), 3);
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Перемещение и изменение размеров фигур

Фигуры располагаются над листом. Их размещение определяется свойством `left` и свойством `top` . Они действуют как поля от соответствующих краев листа, а [0, 0] — верхний левый угол. Их можно задать напрямую или скорректировать из текущей позиции с помощью методов `incrementLeft` и методов `incrementTop` . Величина поворота фигуры от позиции по умолчанию также устанавливается таким образом, `rotation` `incrementRotation` при этом свойством является абсолютная величина, а метод настраивает существующий поворот.

Глубина фигуры относительно других фигур определяется свойством `zorderPosition` . Этот параметр задается с помощью метода `setZOrder` , который принимает [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder` изменяет порядок текущей фигуры относительно других фигур.

Надстройка имеет несколько параметров для изменения высоты и ширины фигур. Установка либо свойства `height` изменяет `width` указанное измерение, не изменяя другое измерение. И `scaleHeight` измените `scaleWidth` соответствующие размеры фигуры относительно текущего или исходного размера (в зависимости от значения [предоставленного ShapeScaleType](/javascript/api/excel/excel.shapescaletype)). [Необязательный параметр ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) указывает, откуда масштабируется фигура (верхний левый угол, средний или правый нижний угол). Если свойство `lockAspectRatio` имеет значение `true`, методы масштабирования поддерживают текущее соотношение аспектов фигуры путем корректировки другого измерения.

> [!NOTE]
> Прямые изменения свойств и `height` свойств `width` влияют только на это свойство, независимо от `lockAspectRatio` его значения.

В следующем примере кода показана фигура, масштабируемая в 1,25 раза больше исходного размера и повернутая на 30 градусов.

```js
// In this sample, the shape "Octagon" is rotated 30 degrees clockwise
// and scaled 25% larger, with the upper-left corner remaining in place.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");

    let shape = sheet.shapes.getItem("Octagon");
    shape.incrementRotation(30);
    shape.lockAspectRatio = true;
    shape.scaleWidth(
        1.25,
        Excel.ShapeScaleType.currentSize,
        Excel.ShapeScaleFrom.scaleFromTopLeft);

    await context.sync();
});
```

## <a name="text-in-shapes"></a>Текст в фигурах

Геометрические фигуры могут содержать текст. Фигуры имеют свойство `textFrame` типа [TextFrame](/javascript/api/excel/excel.textframe). Объект `TextFrame` управляет параметрами отображения текста (например, полями и переполнением текста). `TextFrame.textRange` — это [объект TextRange](/javascript/api/excel/excel.textrange) с текстовым содержимым и параметрами шрифта.

В следующем примере кода создается геометрическая фигура с именем Wave с текстом "Текст фигуры". Он также настраивает цвета фигуры и текста, а также задает выравнивание текста по центру по горизонтали.

```js
// This sample creates a light-blue wave shape and adds the purple text "Shape text" to the center.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let wave = shapes.addGeometricShape(Excel.GeometricShapeType.wave);
    wave.left = 100;
    wave.top = 400;
    wave.height = 50;
    wave.width = 150;

    wave.name = "Wave";
    wave.fill.setSolidColor("lightblue");

    wave.textFrame.textRange.text = "Shape text";
    wave.textFrame.textRange.font.color = "purple";
    wave.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;

    await context.sync();
});
```

Метод `addTextBox` создает тип `ShapeCollection` с `GeometricShape` `Rectangle` белым фоном и черным текстом. Это то же, что и при работе с кнопкой "Текстовое поле" Excel на **вкладке "Вставка**". `addTextBox` принимает строковый аргумент для задания текста .`TextRange`

В следующем примере кода показано создание текстового поля с текстом "Hello!".

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 20;
    textbox.width = 45;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="shape-groups"></a>Группы фигур

Фигуры можно сгруппировать. Это позволяет пользователю рассматривать их как одну сущность для размещения, изменения размера и других связанных задач. [ShapeGroup](/javascript/api/excel/excel.shapegroup) — это тип`Shape`, поэтому ваша надстройка обрабатывает группу как единую фигуру.

В следующем примере кода показаны три фигуры, сгруппированные вместе. В следующем примере кода показано, что группа фигур перемещается вправо на 50 пикселей.

```js
// This sample takes three previously-created shapes ("Square", "Pentagon", and "Octagon")
// and groups them into a single ShapeGroup.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let square = shapes.getItem("Square");
    let pentagon = shapes.getItem("Pentagon");
    let octagon = shapes.getItem("Octagon");

    let shapeGroup = shapes.addGroup([square, pentagon, octagon]);
    shapeGroup.name = "Group";
    console.log("Shapes grouped");

    await context.sync();
});

// This sample moves the previously created shape group to the right by 50 pixels.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shapeGroup = shapes.getItem("Group");
    shapeGroup.incrementLeft(50);
    await context.sync();
});
```

> [!IMPORTANT]
> Отдельные фигуры в группе ссылаются через свойство `ShapeGroup.shapes` , которое имеет тип [GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection). Они больше не доступны через коллекцию фигур листа после группировки. Например, если лист содержит три фигуры и все они сгруппированы, `shapes.getCount` метод листа вернет число 1.

## <a name="export-shapes-as-images"></a>Экспорт фигур в виде изображений

Любой `Shape` объект можно преобразовать в изображение. [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) возвращает строку в кодировке base64. Формат изображения указывается как перечисление [PictureFormat](/javascript/api/excel/excel.pictureformat) , передаваемое в `getAsImage`.

```js
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let shape = shapes.getItem("Image");
    let stringResult = shape.getAsImage(Excel.PictureFormat.png);

    await context.sync();

    console.log(stringResult.value);
    // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
});
```

## <a name="delete-shapes"></a>Удаление фигур

Фигуры удаляются с листа с помощью `Shape` метода `delete` объекта. Другие метаданные не требуются.

В следующем примере кода удаляются все фигуры из **MyWorksheet**.

```js
// This deletes all the shapes from "MyWorksheet".
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("MyWorksheet");
    let shapes = sheet.shapes;

    // We'll load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();

    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    
    await context.sync();
});
```

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Работа с диаграммами с использованием API JavaScript для Excel](excel-add-ins-charts.md)
