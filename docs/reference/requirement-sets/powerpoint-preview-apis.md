---
title: PowerPoint API предварительного просмотра JavaScript
description: Сведения о предстоящих PowerPoint API JavaScript.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 406808b4b4ff16df72d9c37468696525c8be642f
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513993"
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

В следующей таблице перечислены PowerPoint API JavaScript, которые в настоящее время находятся в предварительном просмотре. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. Excel [API JavaScript.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|Указывает, видны ли пули в абзаце.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|Представляет формат пули абзаца.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|Представляет горизонтальное выравнивание абзаца.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|Возвращает формат заливки фигуры.|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|Указывает высоту фигуры в точках.|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|Расстояние в точках от левой стороны фигуры до левой стороны слайда.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|Возвращает формат линии для фигуры.|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|Указывает имя этой фигуры.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|Возвращает объект рамки с текстом для фигуры.|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|Расстояние в точках от верхнего края фигуры до верхнего края слайда.|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|Возвращает тип фигуры.|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|Указывает ширину в точках формы.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|Указывает высоту фигуры в точках.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|Указывает расстояние в точках от левой стороны фигуры до левой стороны слайда.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|Указывает расстояние в точках от верхнего края фигуры до верхнего края слайда.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|Указывает ширину в точках формы.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape (geometricShapeType: PowerPoint. GeometricShapeType, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|Добавляет геометрическую фигуру в слайд.|
||[addLine(connectorType?: PowerPoint. ConnectorType, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|Добавляет строку в слайд.|
||[addTextBox (текст: строка, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|Добавляет текстовое поле на слайд с предоставленным текстом в качестве контента.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|Представляет цвет переднего плана заполнения формы в формате HTML#RRGGBB (например, "FFA500") или в виде htmL-цвета (например, "оранжевый").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|Задает заливку одним цветом для фигуры.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|Указывает процент прозрачности заполнения как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|Возвращает тип заливки фигуры.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|Представление цветового кода HTML текстового цвета (например, "#FF0000" представляет красный цвет).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|Представляет имя шрифта (например, "Калибри").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|Представляет размер шрифта в точках (например, 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|Тип подчеркивания, применяемый для шрифта.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|Представляет цвет строки в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|Представляет стиль тире строки.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|Представляет тип линии фигуры.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|Указывает процент прозрачности строки как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|Указывает, отображается ли форматирование строки элемента фигуры.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|Представляет толщину линии (в пунктах).|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|Автоматические параметры размеров для текстового кадра.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|Удаляет весь текст в рамке с текстом.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|Указывает, содержит ли текстовый кадр текст.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|Представляет вертикальное выравнивание для рамки с текстом.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|Определяет, нарушаются ли строки автоматически, чтобы соответствовать тексту внутри фигуры.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|Возвращает `ShapeFont` объект, который представляет атрибуты шрифта для диапазона текста.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|Возвращает объект `TextRange` для подстройки в заданный диапазон.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|Представляет формат абзаца в текстовом диапазоне.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|Представляет содержимое с обычным текстом в диапазоне текста.|

## <a name="see-also"></a>См. также

- [PowerPoint справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)