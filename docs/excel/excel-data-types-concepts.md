---
title: Основные понятия, связанные с типами данных API JavaScript для Excel
description: Информация об основных понятиях для использования типов данных Excel в надстройках Office.
ms.date: 11/03/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a5d4915638d67c67679095eb03bc04a48e9196dd
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749387"
---
# <a name="excel-data-types-core-concepts-preview"></a>Основные понятия, связанные с типами данных Excel (предварительная версия)

> [!NOTE]
> API типов данных в настоящее время можно использовать только в общедоступной предварительной версии. API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде. Не используйте API предварительной версии в рабочей среде или в важных деловых документах.

> [!IMPORTANT]
> Некоторые описанные в этой статье понятия, связанные с типами данных, например `Range.valuesAsJSON`, находятся в активной разработке, и их общедоступная предварительная версия пока отсутствует. Эта статья служит введением в понятийный аппарат. Описанные в этой статье понятия, которые еще не используются в общедоступной предварительной версии, скоро будут опубликованы в составе этой версии.

В этой статье рассказывается о том, как использовать [API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md) для работы с типами данных. Здесь представлены основные понятия, лежащие в основе разработки типов данных.

## <a name="core-concepts"></a>Основные понятия

Используйте свойство `Range.valuesAsJSON` для работы со значениями типа данных. Это свойство аналогично свойству [Range.values](/javascript/api/excel/excel.range#values), но `Range.values` возвращает только четыре основных типа: значения строки, числа, логического типа или ошибки. Свойство `Range.valuesAsJSON` может возвращать расширенную информацию об этих четырех основных типах, а также такие типы данных, как форматированное число, сущность и веб-изображение.

### <a name="json-schema"></a>Схема JSON

Типы данных используют согласованную схему JSON, которая определяет типы [CellValueType](/javascript/api/excel/excel.cellvaluetype) данных и дополнительные сведения, такие как `basicValue`, `numberFormat` или `address`. Каждый тип `CellValueType` имеет свойства, доступные в соответствии с этим типом. Например, тип `webImage` включает свойства [altText](/javascript/api/excel/excel.webimagecellvalue#altText) и [attribution](/javascript/api/excel/excel.webimagecellvalue#attribution). В следующих разделах приводятся примеры кода JSON для типов форматированного числа, сущности и веб-изображения.

## <a name="formatted-number-values"></a>Форматированное число

Объект [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) позволяет настройкам Excel определять свойство `numberFormat` для некоторого значения. После того как свойство форматированного числа присвоено значению, оно сопровождает это значение в расчетах и может возвращаться функциями.

В следующем примере кода JSON показано значение форматированного числа. Значение форматированного числа `myDate` в примере кода отображается в пользовательском интерфейсе Excel как **1/16/1990**.

```json
// This is an example of the JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    numberFormat: "m/d/yyyy"
};
```

## <a name="entity-values"></a>Сущность

Значение сущности — это контейнер для типов данных, аналогичный объекту в объектно-ориентированном программировании. Сущности также поддерживают массивы в качестве свойств значения сущности. Объект [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) позволяет надстройкам определять такие свойства, как `type`, `text` и `properties`. Свойство `properties` позволяет значению сущности определять и содержать дополнительные типы данных.

В следующем примере кода JSON показано значение сущности, которое содержит текст, изображение, дату и дополнительное текстовое значение.

```json
// This is an example of the JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }
};
```

## <a name="web-image-values"></a>Веб-изображение

Объект [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) создает возможность хранения изображения как части [сущности](#entity-values) или как независимого значения в диапазоне. Этот объект позволяет использовать множество свойств, включая `address`, `altText` и `relatedImagesAddress`.

В следующем примере кода JSON показано, как представлять веб-изображение.

```json
// This is an example of the JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw"
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
- [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Каждый из объектов ошибок может получить доступ к перечисление через свойство `errorSubType`, и в этом перечислении содержатся дополнительные данные об ошибке. Например, объект ошибки `BlockedErrorCellValue` может получить доступ к перечислению [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype). В перечислении `BlockedErrorCellValueSubType` содержатся дополнительные данные о том, что вызвало данную ошибку.

## <a name="see-also"></a>См. также

- [Обзор типов данных в надстройках Excel](excel-data-types-overview.md)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Обзор пользовательских функций и типов данных](custom-functions-data-types-overview.md)