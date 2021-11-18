---
title: Работа с фигурами с PowerPoint API JavaScript
description: Узнайте, как добавлять, удалять и форматирование фигур на PowerPoint слайдах.
ms.date: 10/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: e510ff47f4c54cd465be5c97c5828aad81041c5e
ms.sourcegitcommit: fb4a55764fb60e826ad06d15d1539e41df503b65
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/14/2021
ms.locfileid: "60356865"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api-preview"></a>Работа с фигурами с PowerPoint API JavaScript (предварительный просмотр)

В этой статье описывается использование геометрических фигур, линий и текстовых полей в сочетании с API [Shape](/javascript/api/powerpoint/powerpoint.shape) и [ShapeCollection.](/javascript/api/powerpoint/powerpoint.shapecollection)

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="create-shapes"></a>Создание фигур

Фигуры создаются и хранятся в коллекции фигур слайда ( `slide.shapes` ). `ShapeCollection` имеет несколько `.add*` методов для этой цели. Все фигуры имеют имена и ИД, созданные для них при добавлении в коллекцию. Это и `name` `id` свойства, соответственно. `name` может быть установлено вашей надстройки.

### <a name="geometric-shapes"></a>Геометрические фигуры

Геометрическая фигура создается с одной из перегрузок `ShapeCollection.addGeometricShape` . Первый параметр — это либо [enum GeometricShapeType,](/javascript/api/powerpoint/powerpoint.geometricshapetype) либо строковой эквивалент одного из значений. Существует необязательный второй параметр типа [ShapeAddOptions,](/javascript/api/powerpoint/powerpoint.shapeaddoptions) который может указывать начальный размер фигуры и ее положение относительно верхней и левой сторон слайда, измеряемого в точках. Или эти свойства можно установить после создания фигуры.

В следующем примере кода создается прямоугольник с именем **"Square",** который находится в 100 точках сверху и слева от слайда. Метод возвращает `Shape` объект.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    return context.sync();
});
```

### <a name="lines"></a>Lines

Создается строка с одной из перегрузок `ShapeCollection.addLine` . Первый параметр — это либо [enum ConnectorType,](/javascript/api/powerpoint/powerpoint.connectortype) либо строковой эквивалент одного из значений, чтобы указать, как линия соединится между конечными точками. Существует необязательный второй параметр [типа ShapeAddOptions,](/javascript/api/powerpoint/powerpoint.shapeaddoptions) который может указать точки запуска и окончания строки. Или эти свойства можно установить после создания фигуры. Метод возвращает `Shape` объект.

> [!NOTE]
> Когда фигура является строкой, свойства и свойства объектов и объектов указывают отправную точку линии относительно верхних и левых `top` `left` `Shape` `ShapeAddOptions` краев слайда. Свойства `height` и свойства указывают `width` конечную точку строки *относительно точки начала.* Итак, end point relative to the top and left edges of the slide is `top`  +  `height` () by ( `left`  +  `width` ). Единица измерения для всех свойств — это точки и разрешены отрицательные значения.

В следующем примере кода создается прямая линия на слайде.

```js
// This sample creates a straight line on the first slide.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    return context.sync();
});
```

### <a name="text-boxes"></a>Надписи

Текстовое поле создается [методом addTextBox.](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_) Первый параметр — это текст, который должен сначала отображаться в поле. Существует необязательный второй параметр типа [ShapeAddOptions,](/javascript/api/powerpoint/powerpoint.shapeaddoptions) который может указать начальный размер текстового окна и его положение относительно верхней и левой сторон слайда. Или эти свойства можно установить после создания фигуры.

В следующем примере кода показано, как создать текстовое поле на первом слайде.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    return context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Перемещение и размер фигур

Фигуры сидят поверх слайда. Их размещение определяется `left` свойствами и `top` свойствами. Они выступают в качестве маржи от соответствующих краев слайда, измеряемого в точках, с верхним левом углу и в `left: 0` `top: 0` левом верхнем углу. Размер фигуры определяется свойствами и `height` `width` свойствами. Код может перемещать или реамизировать форму, сбросив эти свойства. (Эти свойства имеют несколько иное значение, когда фигура является строкой. См. [строки.)](#lines)

## <a name="text-in-shapes"></a>Текст в фигурах

Геометрические фигуры могут содержать текст. Фигуры имеют `textFrame` свойство типа [TextFrame](/javascript/api/powerpoint/powerpoint.textframe). Объект `TextFrame` управляет вариантами отображения текста (например, полями и переполнением текста). `TextFrame.textRange` — объект [TextRange](/javascript/api/powerpoint/powerpoint.textrange) с текстовым контентом и настройками шрифтов.

В следующем примере кода создается геометрическая фигура с именем **"Скобки"** с текстом **"Образ текста".** Он также регулирует форму и текстовые цвета, а также задает вертикальное выравнивание текста в центре.

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends and adds the purple text "Shape text" to the center.
PowerPoint.run(function (context) {
    var shapes = context.presentation.slides.getItemAt(0).shapes;
    var braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    return context.sync();
});
```

## <a name="delete-shapes"></a>Удаление фигур

Фигуры удаляются с слайда методом `Shape` `delete` объекта.

В следующем примере кода показано, как удалить фигуры.

```js
PowerPoint.run(function (context) {
    // Delete all shapes from the first slide.
    var sheet = context.presentation.slides.getItemAt(0);
    var shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    return context.sync()
        .then(function () {
            shapes.items.forEach(function (shape) {
                shape.delete()
            });
            return context.sync();
        })
       .catch(errorHandlerFunction);
});
```