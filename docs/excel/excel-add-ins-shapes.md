---
title: Работа с фигурами с Excel API JavaScript
description: Узнайте, Excel определяет фигуры как любой объект, который находится на уровне рисования Excel.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: e035774817c69f7672a2caeb109b9e2706a5efc8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341060"
---
# <a name="work-with-shapes-using-the-excel-javascript-api"></a>Работа с фигурами с Excel API JavaScript

Excel определяет фигуры как любой объект, который находится на уровне рисования Excel. Это означает, что все, что находится за пределами ячейки, — это фигура. В этой статье описывается использование геометрических фигур, линий и изображений в сочетании с API [Shape](/javascript/api/excel/excel.shape) и [ShapeCollection](/javascript/api/excel/excel.shapecollection) . [Диаграммы](/javascript/api/excel/excel.chart) охватываются в своей статье [Work with charts using the Excel API JavaScript](excel-add-ins-charts.md).

На следующем изображении показаны фигуры, которые образуют термометр.
![Изображение термометра, выполненного в виде Excel формы.](../images/excel-shapes.png)

## <a name="create-shapes"></a>Создание фигур

Фигуры создаются и хранятся в коллекции фигур таблицы (`Worksheet.shapes`). `ShapeCollection` имеет несколько `.add*` методов для этой цели. Все фигуры имеют имена и ИД, созданные для них при добавлении в коллекцию. Это и `name` `id` свойства, соответственно. `name` может быть установлено вашей надстройки для легкого получения с помощью `ShapeCollection.getItem(name)` метода.

Следующие типы фигур добавляются с помощью связанного метода.

| Shape | Добавление метода | Подпись |
|-------|------------|-----------|
| Геометрическая фигура | [addGeometricShape](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1)) | `addGeometricShape(geometricShapeType: Excel.GeometricShapeType): Excel.Shape` |
| Изображение (JPEG или PNG) | [addImage](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1)) | `addImage(base64ImageString: string): Excel.Shape` |
| Line | [addLine](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1)) | `addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType): Excel.Shape` |
| SVG | [addSvg](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1)) | `addSvg(xml: string): Excel.Shape` |
| Текстовое поле | [addTextBox](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1)) | `addTextBox(text?: string): Excel.Shape` |

### <a name="geometric-shapes"></a>Геометрические фигуры

Геометрическая фигура создается с `ShapeCollection.addGeometricShape`помощью . Этот метод принимает в качестве аргумента enum [GeometricShapeType](/javascript/api/excel/excel.geometricshapetype) .

В следующем примере кода создается прямоугольник размером 150x150 пикселей с именем **"Square"** , который находится на 100 пикселей с верхней и левой сторон таблицы.

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

Изображения JPEG, PNG и SVG можно вставить в таблицу в форме фигур. В `ShapeCollection.addImage` качестве аргумента метод принимает строку с кодом base64. Это либо образ JPEG или PNG в строковом виде. `ShapeCollection.addSvg` также принимает строку, хотя этот аргумент XML, который определяет графику.

В следующем примере кода показан файл изображения, загружаемый [файлом FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) в качестве строки. Строка имеет метаданные "base64", удалены до создания фигуры.

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

Строка создается с `ShapeCollection.addLine`помощью . Для этого метода необходимы левые и верхние поля точки начала и конца строки. Кроме того, для указания того, как соединителю строки между конечными точками, необходимо также вводить в себя enum [ConnectorType](/javascript/api/excel/excel.connectortype) . В следующем примере кода создается прямая линия на таблице.

```js
// This sample creates a straight line from [200,50] to [300,150] on the worksheet.
await Excel.run(async (context) => {
    let shapes = context.workbook.worksheets.getItem("MyWorksheet").shapes;
    let line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
    line.name = "StraightLine";
    await context.sync();
});
```

Строки можно подключать к другим объектам Shape. Методы `connectBeginShape` и `connectEndShape` начало и окончание строки прикрепляются к фигурам в указанных точках подключения. Расположение этих точек зависит от формы, `Shape.connectionSiteCount` но его можно использовать для обеспечения того, чтобы надстройка не подключалась к точке, не связанной с этим. Строка отключена от любых присоединенных фигур с помощью этих `disconnectBeginShape` и методов `disconnectEndShape` .

В следующем примере кода **строка "MyLine"** соединяется с двумя фигурами с именами **"LeftShape"** **и "RightShape"**.

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

## <a name="move-and-resize-shapes"></a>Перемещение и размер фигур

Фигуры сидят на вершине таблицы. Их размещение определяется свойством и `left` свойством `top` . Они действуют как поля от соответствующих краев таблицы, а [0, 0] — верхний левый угол. Они могут быть установлены непосредственно или скорректированы с текущей позиции с помощью методов `incrementLeft` и методов `incrementTop` . Таким образом устанавливается также размер поворота фигуры из положения по умолчанию, `rotation` `incrementRotation` при этом свойство является абсолютным количеством и методом, корректющим существующее вращение.

Глубина фигуры по отношению к другим фигурам определяется свойством `zorderPosition` . Это устанавливается с помощью метода `setZOrder` , который принимает [ShapeZOrder](/javascript/api/excel/excel.shapezorder). `setZOrder` регулирует порядок текущей фигуры по отношению к другим фигурам.

Надстройка имеет несколько вариантов изменения высоты и ширины фигур. Параметр или свойство `height` изменяет `width` указанное измерение без изменения другого измерения. Соответствующие `scaleHeight` размеры `scaleWidth` фигуры по отношению к текущему или исходному размеру (в зависимости от значения предоставленного [ShapeScaleType](/javascript/api/excel/excel.shapescaletype)) корректируются. [Необязательный параметр ShapeScaleFrom](/javascript/api/excel/excel.shapescalefrom) указывает, откуда масштабирует фигуру (верхний левый угол, средний или нижний правый угол). Если свойство `lockAspectRatio` **верно,** методы масштабирования поддерживают текущее отношение аспектов фигуры, а также корректируют другое измерение.

> [!NOTE]
> Прямые изменения свойств `height` и свойств `width` влияют только на это свойство, `lockAspectRatio` независимо от значения свойства.

В следующем примере кода показана фигура, масштабироваться в 1,25 раза и вращаться на 30 градусов.

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

Геометрические фигуры могут содержать текст. Формы имеют свойство `textFrame` типа [TextFrame](/javascript/api/excel/excel.textframe). Объект `TextFrame` управляет вариантами отображения текста (например, полями и переполнением текста). `TextFrame.textRange` — объект [TextRange](/javascript/api/excel/excel.textrange) с текстовым контентом и настройками шрифтов.

В следующем примере кода создается геометрическая фигура с именем "Wave" с текстом "Shape Text". Он также регулирует форму и текстовые цвета, а также задает горизонтальное выравнивание текста в центре.

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

Метод `addTextBox` создает тип `ShapeCollection` с белым `GeometricShape` `Rectangle` фоном и черным текстом. Это то же самое, что создается Excel на вкладке **Вставки**. `addTextBox` Для закладки текста строки требуется аргумент строки. `TextRange`

В следующем примере кода показано создание текстового окна с текстом "Hello!".

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

## <a name="shape-groups"></a>Группы формы

Фигуры можно сгруппить вместе. Это позволяет пользователю рассматривать их как единую сущность для позиционирования, размеров и других связанных задач. [ShapeGroup](/javascript/api/excel/excel.shapegroup) — это тип`Shape`, поэтому ваша надстройка рассматривает группу как единую фигуру.

В следующем примере кода показаны три фигуры, сгруппироваться вместе. В последующем примере кода показано, что группа фигур перемещается в нужные 50 пикселей.

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
> Отдельные фигуры в группе ссылаются `ShapeGroup.shapes` через свойство, которое имеет тип [GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection). Они больше не доступны в коллекции фигуры таблицы после сгруппии. Например, если ваш таблица имеет три фигуры и все они сгруппировалися, `shapes.getCount` метод таблицы возвращает количество 1.

## <a name="export-shapes-as-images"></a>Экспорт фигур в качестве изображений

Любой `Shape` объект можно преобразовать в изображение. [Shape.getAsImage](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1)) возвращает строку base64-encoded. Формат изображения указывается как переоформаемая [](/javascript/api/excel/excel.pictureformat) `getAsImage`в .

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

Фигуры удаляются из таблицы методом `Shape` объекта `delete` . Другие метаданные не нужны.

В следующем примере кода удаляются все фигуры **из MyWorksheet**.

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
