---
title: Основные понятия, связанные с типами данных API JavaScript для Excel
description: Информация об основных понятиях для использования типов данных Excel в надстройках Office.
ms.date: 09/01/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a5f742e47d698b215a999b966c424819e182ea49
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616023"
---
# <a name="excel-data-types-core-concepts-preview"></a>Основные понятия, связанные с типами данных Excel (предварительная версия)

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

В этой статье рассказывается о том, как использовать [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) для работы с типами данных. Здесь представлены основные понятия, лежащие в основе разработки типов данных.

## <a name="the-valuesasjson-property"></a>Свойство `valuesAsJson`.

Свойство `valuesAsJson` (или единственное для `valueAsJson` [NamedItem](/javascript/api/excel/excel.nameditem)) является неотъемлемой частью создания типов данных в Excel. Это свойство является расширением свойств `values`, таких как [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member). Оба свойства `values` и `valuesAsJson` используются для доступа к значению в ячейке, но свойство `values` возвращает только один из четырех основных типов: строка, число, логическое значение или ошибка (в виде строки). Напротив, свойство `valuesAsJson` возвращает расширенную информацию об этих четырех основных типах, а также такие типы данных, как форматированное число, сущность и веб-изображение.

Следующие объекты предлагают свойство `valuesAsJson`.

- [NamedItem](/javascript/api/excel/excel.nameditem) (как `valueAsJson`)
- [NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)
- [Range](/javascript/api/excel/excel.range)
- [RangeView](/javascript/api/excel/excel.rangeview)
- [TableColumn](/javascript/api/excel/excel.tablecolumn)
- [TableRow](/javascript/api/excel/excel.tablerow)

> [!NOTE]
> Некоторые значения ячеек изменяются в зависимости от языкового стандарта пользователя. Свойство `valuesAsJsonLocal` предлагает поддержку локализации и доступно для всех тех же объектов, что и `valuesAsJson`.

## <a name="cell-values"></a>Значения ячеек

Свойство `valuesAsJson` возвращает псевдоним типа [CellValue](/javascript/api/excel/excel.cellvalue), который является [объединением](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) следующих типов данных.

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

Псевдоним типа `CellValue` также возвращает объект [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties), который является [пересечением](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types) с остальными `*CellValue` типами. Это не тип данных. Свойства объекта `CellValueExtraProperties` используются со всеми типами данных для указания сведений, связанных с перезаписью значений ячеек.

### <a name="json-schema"></a>Схема JSON

Каждый тип значения ячейки, возвращаемый `valuesAsJson`, использует схему метаданных JSON, разработанную для этого типа. Наряду с дополнительными свойствами, уникальными для каждого типа данных, все эти схемы метаданных JSON имеют общие свойства `type`, `basicType`, и `basicValue`

Определяет `type` [cellValueType](/javascript/api/excel/excel.cellvaluetype) данных. Всегда `basicType` доступен только для чтения и используется в качестве резервного, если тип данных не поддерживается или форматируется неправильно. `basicValue` соответствует значению, которое будет возвращено свойством `values`. `basicValue` используется в качестве запасного варианта, когда расчеты сталкиваются с несовместимыми сценариями, такими как более старая версия Excel, не поддерживающая функцию типов данных. Доступен `basicValue` только для чтения для `ArrayCellValue`типов данных , `EntityCellValue`и `LinkedEntityCellValue`типов `WebImageCellValue` данных.

В дополнение к трем полям, общим для всех типов данных, схема метаданных JSON для каждого `*CellValue` имеет свойства, доступные в соответствии с этим типом. Например, тип [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) включает свойства `altText` и `attribution`, а тип [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) предлагает поля `properties` и `text`.

В следующих разделах приводятся примеры кода JSON для типов форматированного числа, сущности и веб-изображения.

## <a name="formatted-number-values"></a>Форматированное число

Объект [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) позволяет настройкам Excel определять свойство `numberFormat` для некоторого значения. После того как свойство форматированного числа присвоено значению, оно сопровождает это значение в расчетах и может возвращаться функциями.

В следующем примере кода JSON показана полная схема значения форматированного числа. Значение форматированного числа `myDate` в примере кода отображается в пользовательском интерфейсе Excel как **1/16/1990**. Если минимальные требования к совместимости для функции типов данных не выполнены, вычисления используют `basicValue` вместо форматированного числа.

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A read-only property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Сущность

Значение сущности — это контейнер для типов данных, аналогичный объекту в объектно-ориентированном программировании. Сущности также поддерживают массивы в качестве свойств значения сущности. Объект [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) позволяет надстройкам определять такие свойства, как `type`, `text` и `properties`. Свойство `properties` позволяет значению сущности определять и содержать дополнительные типы данных.

Свойства `basicType` и `basicValue` определяют, как вычисления читают этот тип данных сущности, если минимальные требования к совместимости для использования типов данных не выполнены. В этом сценарии этот тип данных сущности отображается как ошибка **#VALUE!** в пользовательском интерфейсе Excel.

В следующем примере кода JSON показана полная схема значения сущности, которое содержит текст, изображение, дату и дополнительное текстовое значение.

```TypeScript
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

Значения сущностей также предлагают свойство `layouts`, которое создает карточку для сущности. Карточка отображается в виде модального окна в пользовательском интерфейсе Excel и может демонстрировать дополнительные сведения, содержащиеся в значении сущности, помимо того, что отображается в ячейке. Дополнительные сведения см. в статье [Использование карточек с типами данных значений сущностей](excel-data-types-entity-card.md).

### <a name="linked-entities"></a>Связанные объекты

Значения связанных сущностей или объекты [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) являются типом значения сущности. Эти объекты интегрируют данные, предоставленные внешней службой, и могут отображать эти данные в виде [карты сущности](excel-data-types-entity-card.md), как обычные значения сущности. [Типы данных Stocks и Geography](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877), доступные через пользовательский интерфейс Excel, являются значениями связанных сущностей.

## <a name="web-image-values"></a>Веб-изображение

Объект [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) создает возможность хранения изображения как части [сущности](#entity-values) или как независимого значения в диапазоне. Этот объект позволяет использовать множество свойств, включая `address`, `altText` и `relatedImagesAddress`.

Свойства `basicType` и `basicValue` определяют, как вычисления читают этот тип данных веб-изображения, если минимальные требования к совместимости для использования функции типов данных не выполнены. В этом сценарии этот тип данных веб-изображения отображается как ошибка **#VALUE!** в пользовательском интерфейсе Excel.

В следующем примере кода JSON показана полная схема веб-изображения.

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

## <a name="improved-error-support"></a>Улучшенная поддержка ошибок

API типов данных представляют существующие ошибки пользовательского интерфейса Excel в качестве объектов. Теперь, когда эти ошибки доступны как объекты, надстройки могут определять или извлекать такие свойства, как `type`, `errorType` и `errorSubType`.

Ниже приводится список всех объектов ошибок с поддержкой, расширенной за счет типов данных.

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Каждый из объектов ошибок может получить доступ к перечисление через свойство `errorSubType`, и в этом перечислении содержатся дополнительные данные об ошибке. Например, объект ошибки `BlockedErrorCellValue` может получить доступ к перечислению [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype). В перечислении `BlockedErrorCellValueSubType` содержатся дополнительные данные о том, что вызвало данную ошибку.

## <a name="see-also"></a>См. также

- [Обзор типов данных в надстройках Excel](excel-data-types-overview.md)
- [Использование карточек с типами данных значений сущностей](excel-data-types-entity-card.md)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Пользовательские функции и типы данных](custom-functions-data-types-concepts.md)
