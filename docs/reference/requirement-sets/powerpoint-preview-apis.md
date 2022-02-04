---
title: PowerPoint API предварительного просмотра JavaScript
description: Сведения о предстоящих PowerPoint API JavaScript.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="powerpoint-javascript-preview-apis"></a>PowerPoint API предварительного просмотра JavaScript

Новые PowerPoint API JavaScript сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Управление слайдами | Добавляет поддержку для добавления слайдов, а также управления макетами слайдов и мастерами слайдов. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| Фигуры | Добавляет поддержку для получения ссылок на фигуры на слайде. | [Shape](/javascript/api/powerpoint/powerpoint.shape) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены PowerPoint API JavaScript, которые в настоящее время находятся в предварительном просмотре. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. Excel [API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#powerpoint-powerpoint-bulletformat-visible-member)|Указывает, видны ли пули в абзаце.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-bulletformat-member)|Представляет формат пули абзаца.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-horizontalalignment-member)|Представляет горизонтальное выравнивание абзаца.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-fill-member)|Возвращает формат заливки фигуры.|
||[height](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-height-member)|Указывает высоту фигуры в точках.|
||[left](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-left-member)|Расстояние в точках от левой стороны фигуры до левой стороны слайда.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-lineformat-member)|Возвращает формат линии для фигуры.|
||[name](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-name-member)|Указывает имя этой фигуры.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-textframe-member)|Возвращает объект рамки с текстом для фигуры.|
||[top](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-top-member)|Расстояние в точках от верхнего края фигуры до верхнего края слайда.|
||[type](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-type-member)|Возвращает тип фигуры.|
||[width](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-width-member)|Указывает ширину в точках формы.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-height-member)|Указывает высоту фигуры в точках.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-left-member)|Указывает расстояние в точках от левой стороны фигуры до левой стороны слайда.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-top-member)|Указывает расстояние в точках от верхнего края фигуры до верхнего края слайда.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-width-member)|Указывает ширину в точках формы.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape (geometricShapeType: PowerPoint. GeometricShapeType, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgeometricshape-member(1))|Добавляет геометрическую фигуру в слайд.|
||[addLine(connectorType?: PowerPoint. ConnectorType, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addline-member(1))|Добавляет строку в слайд.|
||[addTextBox (текст: строка, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1))|Добавляет текстовое поле на слайд с предоставленным текстом в качестве контента.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-clear-member(1))|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-foregroundcolor-member)|Представляет цвет переднего плана заполнения формы в формате HTML#RRGGBB (например, "FFA500") или в виде htmL-цвета (например, "оранжевый").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-setsolidcolor-member(1))|Задает заливку одним цветом для фигуры.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-transparency-member)|Указывает процент прозрачности заполнения как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-type-member)|Возвращает тип заливки фигуры.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-bold-member)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-color-member)|Представление цветового кода HTML текстового цвета (например, "#FF0000" представляет красный цвет).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-italic-member)|Указывает, применяется ли курсив.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-name-member)|Представляет имя шрифта (например, "Калибри").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-size-member)|Представляет размер шрифта в точках (например, 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-underline-member)|Тип подчеркивания, применяемый для шрифта.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-color-member)|Представляет цвет строки в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-dashstyle-member)|Представляет стиль тире строки.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-style-member)|Представляет тип линии фигуры.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-transparency-member)|Указывает процент прозрачности строки как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-visible-member)|Указывает, отображается ли форматирование строки элемента фигуры.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-weight-member)|Представляет толщину линии (в пунктах).|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-autosizesetting-member)|Автоматические параметры размеров для текстового кадра.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-bottommargin-member)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-deletetext-member(1))|Удаляет весь текст в рамке с текстом.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-hastext-member)|Указывает, содержит ли текстовый кадр текст.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-leftmargin-member)|Представляет левое поле рамки с текстом (в пунктах).|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-rightmargin-member)|Представляет правое поле рамки с текстом (в пунктах).|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-textrange-member)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-topmargin-member)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-verticalalignment-member)|Представляет вертикальное выравнивание для рамки с текстом.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-wordwrap-member)|Определяет, нарушаются ли строки автоматически, чтобы соответствовать тексту внутри фигуры.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-font-member)|Возвращает объект `ShapeFont` , который представляет атрибуты шрифта для диапазона текста.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-getsubstring-member(1))|Возвращает объект `TextRange` для подстройки в заданный диапазон.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-paragraphformat-member)|Представляет формат абзаца в текстовом диапазоне.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-text-member)|Представляет содержимое с обычным текстом в диапазоне текста.|

## <a name="see-also"></a>См. также

- [PowerPoint справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)