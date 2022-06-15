---
title: Работа с фигурами с помощью PowerPoint API JavaScript
description: Узнайте, как добавлять, удалять и форматировать фигуры на PowerPoint слайдах.
ms.date: 06/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f314cfebb26450e79dbabe1e65ac9e4c8fe9799
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091106"
---
# <a name="work-with-shapes-using-the-powerpoint-javascript-api"></a>Работа с фигурами с помощью PowerPoint API JavaScript

В этой статье описывается, как использовать геометрические фигуры, линии и текстовые поля в сочетании с [API-интерфейсами Shape и ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection).[](/javascript/api/powerpoint/powerpoint.shape)

## <a name="create-shapes"></a>Создание фигур

Фигуры создаются и хранятся в коллекции фигур слайда ().`slide.shapes` `ShapeCollection` имеет несколько `.add*` методов для этой цели. Все фигуры имеют имена и идентификаторы, созданные для них при добавлении в коллекцию. Это свойства `name` `id` и свойства соответственно. `name` может быть задано надстройка.

### <a name="geometric-shapes"></a>Геометрические фигуры

Геометрическая фигура создается с одной из перегрузок `ShapeCollection.addGeometricShape`. Первый параметр — это перечисление [GeometrShapeType](/javascript/api/powerpoint/powerpoint.geometricshapetype) или строка, эквивалентная одному из значений перечисления. Существует необязательный второй параметр типа [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) , который может указывать начальный размер фигуры и ее положение относительно верхней и левой сторон слайда, измеряемого в точках. Или эти свойства можно задать после создания фигуры.

В следующем примере кода создается прямоугольник с именем **Square** , который располагается на 100 точек сверху и слева от слайда. Метод возвращает `Shape` объект.

```js
// This sample creates a rectangle positioned 100 points from the top and left sides
// of the slide and is 150x150 points. The shape is put on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    rectangle.left = 100;
    rectangle.top = 100;
    rectangle.height = 150;
    rectangle.width = 150;
    rectangle.name = "Square";
    await context.sync();
});
```

### <a name="lines"></a>Lines

Создается строка с одной из перегрузок `ShapeCollection.addLine`. Первый параметр — это либо перечисление [ConnectorType](/javascript/api/powerpoint/powerpoint.connectortype) , либо строка, эквивалентная одному из значений перечисления, чтобы указать, как строка соединяется между конечными точками. Существует необязательный второй параметр типа [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) , который может указывать конечную и конечную точки строки. Или эти свойства можно задать после создания фигуры. Метод возвращает `Shape` объект.

> [!NOTE]
> Если фигура является линией, `top` `left` `Shape` `ShapeAddOptions` свойства и объекты указывают начальную точку линии относительно верхнего и левого краев слайда. И `height` свойства `width` указывают конечную точку строки *относительно начальной точки*. Таким образом, конечная точка относительно верхнего и левого краев слайда будет (`top` + `height`) по (`left` + `width`). Единицей измерения для всех свойств являются точки, а отрицательные значения разрешены.

В следующем примере кода создается прямая линия на слайде.

```js
// This sample creates a straight line on the first slide.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(Excel.ConnectorType.straight, {left: 200, top: 50, height: 300, width: 150});
    line.name = "StraightLine";
    await context.sync();
});
```

### <a name="text-boxes"></a>Надписи

Текстовое поле создается с помощью [метода addTextBox](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1)) . Первый параметр — это текст, который должен изначально отображаться в поле. Существует необязательный второй параметр типа [ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions) , который может указывать начальный размер текстового поля и его положение относительно верхней и левой сторон слайда. Или эти свойства можно задать после создания фигуры.

В следующем примере кода показано, как создать текстовое поле на первом слайде.

```js
// This sample creates a text box with the text "Hello!" and sizes it appropriately.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox("Hello!");
    textbox.left = 100;
    textbox.top = 100;
    textbox.height = 300;
    textbox.width = 450;
    textbox.name = "Textbox";
    await context.sync();
});
```

## <a name="move-and-resize-shapes"></a>Перемещение и изменение размеров фигур

Фигуры находятся в верхней части слайда. Их размещение определяется свойствами `left` и свойствами `top` . Они действуют как поля от соответствующих краев слайда, измеряемые в точках, `left: 0` `top: 0` с верхним левым углом и справа от них. Размер фигуры определяется свойствами `height` и свойствами `width` . Код может перемещать или изменить размер фигуры, сбрасывая эти свойства. (Эти свойства имеют немного другое значение, если фигура является линией. См. [строки](#lines).)

## <a name="text-in-shapes"></a>Текст в фигурах

Геометрические фигуры могут содержать текст. Фигуры имеют свойство `textFrame` типа [TextFrame](/javascript/api/powerpoint/powerpoint.textframe). Объект `TextFrame` управляет параметрами отображения текста (например, полями и переполнением текста). `TextFrame.textRange` — это [объект TextRange](/javascript/api/powerpoint/powerpoint.textrange) с текстовым содержимым и параметрами шрифта.

В следующем примере кода создается геометрическая фигура с именем **"Фигурные** скобки" с текстом **"Текст фигуры"**. Он также изменяет цвета фигуры и текста, а также устанавливает вертикальное выравнивание текста по центру.

```js
// This sample creates a light blue rectangle with braces ("{}") on the left and right ends
// and adds the purple text "Shape text" to the center.
await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    await context.sync();
});
```

## <a name="delete-shapes"></a>Удаление фигур

Фигуры удаляются со слайда с помощью `Shape` метода `delete` объекта.

В следующем примере кода показано, как удалить фигуры.

```js
await PowerPoint.run(async (context) => {
    // Delete all shapes from the first slide.
    const sheet = context.presentation.slides.getItemAt(0);
    const shapes = sheet.shapes;

    // Load all the shapes in the collection without loading their properties.
    shapes.load("items/$none");
    await context.sync();
        
    shapes.items.forEach(function (shape) {
        shape.delete();
    });
    await context.sync();
});
```
